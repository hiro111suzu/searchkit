//. init
var actx = {};
_common_lib();
var tab_name = getTabName( GetCurrentTab() ).trim();
var prefix_from_tab = tab_name.split_l(' ');
var term_from_tab   = tab_name.split_r(' ');

//.. �ݒ�
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
_debug_log( func, '�֐��̎��s��' );
eval( func );
endMacro();

//. func �R�}���h

//.. es_search
function es_search( term ) {
	term = term || _input_term( 'Everything����' );
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
	term = term || _input_term( '�����ύX�����t�@�C������' );
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
	term = term || _input_term( '�T�u�t�H���_�[���̌���\n' + user_conf.set_dir );
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
	term = term || _input_term( 'Windows Index����' );
	// ����
	var ado_conn = new ActiveXObject("ADODB.Connection");
	var ado_rec  = new ActiveXObject("ADODB.Recordset");
	ado_conn.Open( "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';" );
	ado_rec.Open(
		"SELECT System.ItemUrl FROM SYSTEMINDEX WHERE Contains('" + term + "')" ,
		ado_conn
	);
	// �����o��
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
		message( '�^�u�����猟���@�����m�ł��܂���' );
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
		//- �N���b�v�t�H���_�[
		_exec_script( 'clipfol', 'reload' )
	} else if ( tab_name == '*�G�ۃq�X�g��' ) {
		//- �G�ۃq�X�g��
		closeTab( GetCurrentTab() );
		_run( 'hidemaru', ' /m3 /x hidehist\\hidehist.mac' );
	} else {
		Command( '�t�B���^��S�ĉ���' );
		Command( '�ŐV�̏��ɍX�V' );
	}
}

//.. filter
function filter( arg ) {
	term = arg || _input_term( '�Č���', term_from_tab );
	if ( term && term.charAt(0) == '/' ) {
		selectItem( term.substring(1).add_wildcard() );
		setSpaceSelectMode(1);
	} else {
		refresh( term )
	}
}

//.. filter_plus �t�B���^�[�E�Č���
function filter_plus( arg ) {
	//... �����^�u
	if ( tab_name.charAt(0) == '?' ) {
		filter( arg );
		return;
	}

	//... �ʏ�^�u
	var filt = arg ||
		input( '�t�B���^������', GetWildcard().replace( /^\*(.+)\*$/, '$1' ) );

	var flg_slash = false;
	if ( filt && filt.charAt(0) == '/' ) {
		flg_slash = true;
		filt = filt.substring(1);
	}
	if ( ! flg_slash ) {
		if ( !filt ) {
			Command( '�t�B���^��S�ĉ���' );
			return;
		}
		Open( filt.add_wildcard() );
	} else {
		if ( !filt ) return;
		selectItem( filt.add_wildcard() );
		setSpaceSelectMode(1);
	}
}

//. func �c�[��
//.. _custom_fl �J�X�^���t�@�C�����X�g�쐬
function _custom_fl( term ) {
	if ( ! user_conf.tab_reuse )
		Command( '�V�����^�u' );

	var attention = '��������: ' + term + ' | �q�b�g��: ';
	var fn = user_conf.fn_listtext;
	if ( actx.fs.GetFile( fn ).Size ) {
		var n = actx.fs.OpenTextFile( fn, 8 ).Line - 1;
		attention += n == user_conf.max_num
			? user_conf.max_num + '���ȏ�'
			: n + '��'
		;
	} else {
		attention += '�Ȃ�';
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
		message( '�s���Ȍ����@\n�Č����R�}���h' );
	}
}

//.. _run_search
function _run_search() {
	var cmd = Array.prototype.join.call( arguments, ' ' ); //- �X�y�[�X��؂�őS���Ȃ���
	_debug_log( cmd, '�R�}���h���C��' );
	actx.shell.run(
		cmd ,
		0, // hide
		true // wait'
	);
}

//.. _input_term
function _input_term( msg, pre ) {
	var ret = input( msg + '\n����������', pre );
	if ( ret )
		return ret;
	else
		endMacro();
}


//.. _save_data
function _save_data( data, file_name ) {
	obj_ado = new ActiveXObject( "ADODB.Stream" );
	obj_ado.Type = 2;	// -1:Binary, 2:Text
	obj_ado.Mode = 3;	// �ǂݎ��/�������݃��[�h
	obj_ado.charset = 'UTF-8';
	obj_ado.LineSeparator = -1;  // ' -1 CrLf , 10 Lf , 13 Cr
	obj_ado.Open();
	obj_ado.WriteText( data.join( user_conf.new_line ) + user_conf.new_line, 0 );
	obj_ado.SaveToFile( file_name || user_conf.fn_listtext, 2 );
}

//. func �ėp
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
			+ fn_script.q() //- script ��
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
	if ( ! question( str + '\n\n�u�������v�ŏI��' ) ) endMacro();
}

//.. _debug_log
function _debug_log( val, key ) {
	if ( ! user_conf || ! user_conf.debug_soft ) return;
	if ( ! val ) { //- �J�n
		if ( ! user_conf.debug_soft.is_file() ) {
			message( 'debug_soft�̎��s�t�@�C��������܂���\n' + user_conf.debug_soft  );
			user_conf.debug_soft = null;
			return;
		}
		val = '�J�n';
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

		//... string �g��
	//- �e�f�B���N�g��
	String.prototype.parent = function(){
		return actx.fs.GetParentFolderName( this );
	}

	//- basename
	String.prototype.basename = function(){
		return actx.fs.GetBaseName( this );
	}

	//- �t���p�X��
	String.prototype.fullpath = function( ext ){
		return GetDirectory() + '\\' + this + ( ext ? '.' + ext : '' );
	}

	//- �g���q
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

	//- �_�u���N�I�[�e�[�V�����ň͂�
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
