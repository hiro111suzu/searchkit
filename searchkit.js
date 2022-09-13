//. init
var actx = {};
_common_lib();
var tab_name = getTabName( GetCurrentTab() ).trim();
var prefix_from_tab = tab_name.split_l(' ');
var term_from_tab   = tab_name.split_r(' ');

//.. 設定
var user_conf = {
	max_num: 1000,
	prefix2func: {
		'?e': 'es_search' ,
		'?w': 'win_search' ,
		'?today': 'es_today' ,
		'?sub': 'es_subfolder' ,
	} ,
	default_search: 'es_search' ,
	fn_listtext: '%TEMP%\\search.txt'.env_expand() ,
	set_dir: 'c:\\$' ,
	es_exe: scriptFullName.parent() + '\\es.exe' ,
	tab_reuse: false ,
//	debug_soft: '...\\HmSharedOutputPane_x64\\HmSharedOutputPane.exe' ,
	new_line: '\r\n'
}
_debug_log();

//. main
var arg = [ getArg(0), getArg(1) ], func = '', term = '';
_debug_log( arg, 'arg' );
var func = user_conf.default_search;
if ( arg[0] ) {
	if ( eval( 'typeof ' + arg[0] ) == 'function' ) {
		func = arg[0];
		term = arg[1];
	} else {
		term = arg[0];
	}
}

var prefix = _func2prefix( func );
func += '( term )';
_debug_log( func, '関数の実行文' );
eval( func );
endMacro();

//. func コマンド

//.. es_search
function es_search( term ) {
	term = term || _input_term( 'Everything検索' );
	_run_search(
		user_conf.es_exe.q() ,
		'-max-results', user_conf.max_num,
		'-export-txt', user_conf.fn_listtext.q(),
		term
	);
	_custom_fl( term );
}

//.. es_today
function es_today( term ) {
	term = term || _input_term( '今日変更したファイル検索' );
	_run_search(
		user_conf.es_exe.q() ,
		'-max-results', user_conf.max_num,
		'-export-txt', user_conf.fn_listtext.q(),
		'recentchange:today' ,
		term
	);
	_custom_fl( term );
}

//.. es_subfolder
function es_subfolder( term ) {
	user_conf.set_dir = getDirectory();
	term = term || _input_term( 'サブフォルダー内の検索\n' + user_conf.set_dir );
	_run_search(
		user_conf.es_exe.q() ,
		'-max-results', user_conf.max_num,
		'-export-txt', user_conf.fn_listtext.q(),
		'-path', user_conf.set_dir.q() ,
		term
	);
	_custom_fl( term );
}

//.. win_search
function win_search( term ) {
	term = term || _input_term( 'Windows Index検索' );
	// 検索
	var ado_conn = new ActiveXObject("ADODB.Connection");
	var ado_rec  = new ActiveXObject("ADODB.Recordset");
	ado_conn.Open( "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';" );
	ado_rec.Open(
		"SELECT System.ItemUrl FROM SYSTEMINDEX WHERE Contains('" + term + "')" ,
		ado_conn
	);
	// 書き出し
	var list = [];
	while ( ! ado_rec.EOF ) {
		list.push( String( ado_rec.Fields.Item("System.ItemUrl") ) );
		ado_rec.MoveNext();
	}
	_save_data( list );
	_custom_fl( term );
}

//.. refresh
function refresh( term ) {
	term = term || term_from_tab;
	if ( ! term ) return;
	prefix = prefix_from_tab;
	var func = user_conf.prefix2func[ prefix ];
	if ( ! func ) {
		message( 'タブ名から検索法を検知できません' );
		return;
	}
	user_conf.tab_reuse = true;
	eval( func + '(term)' );
}

//.. refresh_plus
function refresh_plus() {
	var tab_name0 = tab_name.charAt(0);
	if ( tab_name0 == '?' ) {
	   refresh();
	} else if ( tab_name0 == '*' ) {
		//- クリップフォルダー
		_exec_script( 'clipfol', 'reload' )
	} else if ( tab_name == '*秀丸ヒストリ' ) {
		//- 秀丸ヒストリ
		closeTab( GetCurrentTab() );
		_run( 'hidemaru', ' /m3 /x hidehist\\hidehist.mac' );
	} else {
		Command( 'フィルタを全て解除' );
		Command( '最新の情報に更新' );
	}
}

//.. filter
function filter( arg ) {
	term = arg || _input_term( '再検索', term_from_tab );
	if ( term && term.charAt(0) == '/' ) {
		selectItem( term.substring(1).add_wildcard() );
		setSpaceSelectMode(1);
	} else {
		refresh( term )
	}
}

//.. filter_plus フィルター・再検索
function filter_plus( arg ) {
	//... 検索タブ
	if ( tab_name.charAt(0) == '?' ) {
		filter( arg );
		return;
	}

	//... 通常タブ
	var filt = arg ||
		input( 'フィルタ文字列', GetWildcard().replace( /^\*(.+)\*$/, '$1' ) );

	var flg_slash = false;
	if ( filt && filt.charAt(0) == '/' ) {
		flg_slash = true;
		filt = filt.substring(1);
	}
	if ( ! flg_slash ) {
		if ( !filt ) {
			Command( 'フィルタを全て解除' );
			return;
		}
		Open( filt.add_wildcard() );
	} else {
		if ( !filt ) return;
		selectItem( filt.add_wildcard() );
		setSpaceSelectMode(1);
	}
}

//. func ツール
//.. _custom_fl カスタムファイルリスト作成
function _custom_fl( term ) {
	if ( ! user_conf.tab_reuse )
		Command( '新しいタブ' );

	var attention = '検索条件: ' + term + ' | ヒット数: ';
	var fn = user_conf.fn_listtext;
	if ( actx.fs.GetFile( fn ).Size ) {
		var n = actx.fs.OpenTextFile( fn, 8 ).Line - 1;
		attention += n == user_conf.max_num
			? user_conf.max_num + '件以上'
			: n + '件'
		;
	} else {
		attention += 'なし';
	}
	setAttentionBar( attention );
	customFileList( fn );
	setTabName( prefix + ' ' + term, GetCurrentTab() );
	setDirectory( user_conf.set_dir );
	setAttentionBar( attention );
}

//.. _func2prefix
function _func2prefix( func ) {
	for ( var key in user_conf.prefix2func ) {
		if ( func != user_conf.prefix2func[ key ] ) continue;
		return key;
	}
	return '';
}

//.. _redo
function _redo( type, term ) {
	if ( ! term || ! type ) return;
	user_conf.tab_reuse = true;
	if ( type == 'e' ) {
		es_search( term );
	} else if ( type == 'w' ) {
		win_search( term );
	} else {
		message( '不明な検索法\n再検索コマンド' );
	}
}

//.. _run_search
function _run_search() {
	var cmd = Array.prototype.join.call( arguments, ' ' ); //- スペース区切りで全部つなげる
	_debug_log( cmd, 'コマンドライン' );
	actx.shell.run(
		cmd ,
		0, // hide
		true // wait'
	);
}

//.. _input_term
function _input_term( msg, pre ) {
	var ret = input( msg + '\n検索語を入力', pre );
	if ( ret )
		return ret;
	else
		endMacro();
}


//.. _save_data
function _save_data( data, file_name ) {
	obj_ado = new ActiveXObject( "ADODB.Stream" );
	obj_ado.Type = 2;	// -1:Binary, 2:Text
	obj_ado.Mode = 3;	// 読み取り/書き込みモード
	obj_ado.charset = 'UTF-8';
	obj_ado.LineSeparator = -1;  // ' -1 CrLf , 10 Lf , 13 Cr
	obj_ado.Open();
	obj_ado.WriteText( data.join( user_conf.new_line ) + user_conf.new_line, 0 );
	obj_ado.SaveToFile( file_name || user_conf.fn_listtext, 2 );
}

//. func 汎用
//.. _exec_script
function _exec_script( name, args ) {
	var s = name.env_expand();
	var p = scriptFullName.parent();
	var path_set = [
		s ,
		p + '\\' + s ,
		p.parent() + '\\' + s ,
		p.parent() + '\\' + s + '\\' + s
	];
	var ext_set = [ '', '.js', '.vbs' ];
	for ( var num in path_set ) for ( var num2 in ext_set ) {
		var fn_script = path_set[ num ] + ext_set[ num2 ];
		if ( ! fn_script.is_file() ) continue;
		sleep( 50 );
		actx.shell.run( FullName, "/m3 /x "
			+ fn_script.q() //- script 名
			+ ( args && ' /a ' + ( typeof args == "string"
				? args
				: args.join( ' /a ' )
			))
		);
		return true;
	}
}

//.. _test
function _test( str ) {
	if ( ! question( str + '\n\n「いいえ」で終了' ) ) endMacro();
}

//.. _debug_log
function _debug_log( val, key ) {
	if ( ! user_conf || ! user_conf.debug_soft ) return;
	if ( ! val ) { //- 開始
		if ( ! user_conf.debug_soft.is_file() ) {
			message( 'debug_softの実行ファイルがありません\n' + user_conf.debug_soft  );
			user_conf.debug_soft = null;
			return;
		}
		val = '開始';
	}
	var type = typeof val;
	actx.shell.run(
		user_conf.debug_soft + ' '
		+ (
			'[' + scriptFullName.basename() + '] '
			+ ( key ? key +  '(' + type + ')\r\n' : '' )
			+ ( type == 'object' ? _JSON_encode( val ) : val )
		).q() ,
		0 //- hide
	);
}

//.. _JSON_encode
function _JSON_encode( obj ){
	var htmlfile = new ActiveXObject( 'htmlfile' );
	htmlfile.write( '<meta http-equiv="x-ua-compatible" content="IE=11">' );
	return htmlfile.parentWindow.JSON.stringify( obj );
}

	//.. _common_lib
function _common_lib() {
	actx = {
		fs: new ActiveXObject( "Scripting.FileSystemObject" ) ,
		shell: new ActiveXObject( "WScript.Shell" )
	};

		//... string 拡張
	//- 親ディレクトリ
	String.prototype.parent = function(){
		return actx.fs.GetParentFolderName( this );
	}

	//- basename
	String.prototype.basename = function(){
		return actx.fs.GetBaseName( this );
	}

	//- フルパスに
	String.prototype.fullpath = function( ext ){
		return GetDirectory() + '\\' + this + ( ext ? '.' + ext : '' );
	}

	//- 拡張子
	String.prototype.ext = function(){
		return actx.fs.GetExtensionName( this );
	}

	// is_file
	String.prototype.is_file = function(){
		return actx.fs.FileExists( this );
	}

	// is_folder
	String.prototype.is_folder = function(){
		return actx.fs.FolderExists( this );
	}

	//- ダブルクオーテーションで囲む
	String.prototype.q = function(){
		return '"' + this.replace( /"/g, '""' )  + '"';
	}
	//- has
	String.prototype.has = function( str ){
		return this.indexOf( str ) != -1;
	}
	//- add_wildcard
	String.prototype.add_wildcard = function() {
		return this.has( '*' ) ? this : '*' + this + '*' 
	}
	//- trim
	String.prototype.trim = function() {
		return this.replace( /^\s+|\s+$/g, '' );
	}
	//- env_expand
	String.prototype.env_expand = function() {
		return actx.shell.ExpandEnvironmentStrings( this );
	}
	String.prototype.split_l = function( sep ) {
		return this.split( sep, 2 )[0];
	}
	String.prototype.split_r = function( sep ) {
		return this.substring( this.split( sep, 2 )[0].length + sep.length );
	}
}
