function BitProxyRemote(){
	var logged_in = false, password = '';
	
	function tabClicked(id){
		if (!logged_in) {
		}
		
		$('#topmenu .tab').removeClass('active');
		$('#tab_'+id).addClass('active');
	}
	
	function makeAjaxRequest(func, params, callback){
		$('#ajax_loader').show();
		
		$.post(
			'http://127.0.0.1/btp/'+password+'/'+func, 
			params,
			function(data){
				$('#ajax_loader').hide();
				if (callback)
					callback(data);
			},
			'json'
		);
	}
	
	function showError(msg){
		
	}
	
	function init(){
		$('#topmenu .tab').click(function(){
			tabClicked($(this).attr('data-id'));
			return false;
		});
		$('#auth-do').click(function(){
			var pass = $('#auth-pass').val();
			if (pass == '') {
				showError('Пароль не может быть пустым');
				return;
			}
			password = pass;
			makeAjaxRequest('login', {}, function(data){
				
			});
		});
	}
	
	init();
}
$(function(){
	BitProxyRemote();
});


var logged_in = false;
var adradd = '';
var cur_pass = '';

var got_downloading = false;
var got_emulation = false;
var got_info = false;

var emul_inner_view = '';
var down_inner_view = '';

function callback_error(err_num, desc)
{
	switch (err_num)
	{
		case 1:
			err = 'Ошибка в протоколе';
			break;
		case 2:
			err = 'Не верный пароль';
			break;
		case 3:
			err = 'Не известное действие';
			break;
		case 999:
			err = desc;
			break;
		default:
			err = 'Не известная ошибка';
			break;
	}
	show_error(err);
}

function callback_logged()
{
	logged_in = true;
	cur_pass = document.getElementById('my_pass').value;
	document.getElementById('login_box').style.display = 'none';
}

function callback_info(ver, time, down, emul)
{
	highlight_menu(3);
	got_info = true;

	document.getElementById('infobox_ver').innerHTML = ver;
	document.getElementById('infobox_time').innerHTML = time;
	document.getElementById('infobox_down').innerHTML = down;
	document.getElementById('infobox_emul').innerHTML = emul;

	hide_main();
	document.getElementById('info_box').style.display = '';
}

function callback_emul_add(id, tname, tracker, status, uploaded, downloaded, complete, reiting, peers)
{
	emul_inner_view += '<tr class="normal" onmouseover="this.className=\'normal_hover\'" onmouseout="this.className=\'normal\'">';
	emul_inner_view += '<td align="left">' + tname + '<br /><span class="small_text">Трекер: ' + tracker + '</span></td>';
	emul_inner_view += '<td align="center">' + status + '<br />Залито: ' + uploaded + '<br />Скачано: ' + downloaded + '<br />';
	emul_inner_view += 'Завершено: ' + complete + '<br />Рейтинг: ' + reiting + '</td>';
	emul_inner_view += '<td align="center">' + peers + '</td>';
	emul_inner_view += '<td><a href="javascript:;" onClick="do_nop()">Запустить</a><br /><a href="javascript:;" onClick="do_nop()">Остановить</a><br /><a href="javascript:;" onClick="do_nop()">Удалить</a><br /><a href="javascript:;" onClick="do_nop()">Параметры</a></td>';
	emul_inner_view += '</tr>';
}

function callback_down_add(id, tname, tracker, port, mode, ison)
{
	down_inner_view += '<tr class="normal" onmouseover="this.className=\'normal_hover\'" onmouseout="this.className=\'normal\'">';
	down_inner_view += '<td align="left">' + tname + '<br /><span class="small_text">Трекер: ' + tracker + '</span></td>';
	down_inner_view += '<td align="center">' + ((ison == '1') ? 'Включен' : 'Выключен') + '<br />Режим: ' + mode + '<br />';
	down_inner_view += 'Порт: ' + port + '</td>';
	down_inner_view += '<td><a href="javascript:;" onClick="do_nop()">Включить</a><br /><a href="javascript:;" onClick="do_nop()">Отключить</a><br /><a href="javascript:;" onClick="do_nop()">Удалить</a><br /><a href="javascript:;" onClick="do_nop()">Параметры</a></td>';
	down_inner_view += '</tr>';
}

function callback_emul_flash()
{
	highlight_menu(2);
	got_emulation = true;
	emul_inner_view += '<tr class="table_head"><td colspan="4" align="center"><a href="javascript:;" onClick="do_nop()">Добавить задание</a></td></tr>'
	document.getElementById('emulation_box').innerHTML = emul_inner_view;

	hide_main();
	document.getElementById('emulation_box').style.display = '';
}

function callback_down_flash()
{
	highlight_menu(1);
	got_downloading = true;
	document.getElementById('download_box').innerHTML = down_inner_view;

	hide_main();
	document.getElementById('download_box').style.display = '';
}

function hide_main()
{
	document.getElementById('download_box').style.display = 'none';
	document.getElementById('emulation_box').style.display = 'none';
	document.getElementById('info_box').style.display = 'none';
}

function show_loginbox()
{
	document.getElementById('login_box').style.display = '';
}

function highlight_menu(j)
{
	for (var i = 1; i <= 3; i++)
	{
		document.getElementById('tab_' + i).className = (i == j) ? 'tabactive' : 'tab';
	}
}

function reset_error()
{
	document.getElementById('error_box').style.display = 'none';
}

function reset_emulation()
{
	emul_inner_view = '<tr class="table_head"><td colspan="4" align="center">Эмуляция (<a href="javascript:;" onClick="request_emulation()">обновить</a>) </td></tr>';
	emul_inner_view += '<tr align="center" class="table_head"><td>Раздача</td><td>Состояние</td><td>P/S</td><td>Опции</td></tr>';
}

function reset_download()
{
	down_inner_view = '<tr class="table_head"><td colspan="3" align="center">Скачка (<a href="javascript:;" onClick="request_download()">обновить</a>) </td></tr>';
	down_inner_view += '<tr align="center" class="table_head"><td>Раздача</td><td>Состояние</td><td>Опции</td></tr>';
}

function show_error(desc)
{
	document.getElementById('error_desc').innerHTML = desc;
	document.getElementById('error_box').style.display = '';
}

function try_login()
{
	reset_error();

	var ajax = new iAjax();
	var pass = document.getElementById('my_pass').value;
	if (pass == '')
	{
		alert('Пароль не указан!');
		return;
	}
	ajax.execute = true;
	ajax.requestFile = adradd + "/btp/" + pass + '/login';
	ajax.runAJAX();
}

function request_info()
{
	var ajax = new iAjax();
	ajax.execute = true;
	ajax.requestFile = adradd + "/btp/" + cur_pass + '/info';
	ajax.runAJAX();
}

function request_emulation()
{
	var ajax = new iAjax();
	ajax.execute = true;
	ajax.requestFile = adradd + "/btp/" + cur_pass + '/emul';
	ajax.runAJAX();
}

function request_download()
{
	var ajax = new iAjax();
	ajax.execute = true;
	ajax.requestFile = adradd + "/btp/" + cur_pass + '/down';
	ajax.runAJAX();
}

function show_info()
{
	if (!logged_in)
	{
		show_loginbox();
		return;
	}
	if (got_info)
	{
		highlight_menu(3);
		hide_main();
		document.getElementById('info_box').style.display = '';
	} else {
		request_info();
	}
}

function show_emulation()
{
	if (!logged_in)
	{
		show_loginbox();
		return;
	}
	if (got_emulation)
	{
		highlight_menu(2);
		hide_main();
		document.getElementById('emulation_box').style.display = '';
	} else {
		request_emulation();
	}
}

function show_downloading()
{
	if (!logged_in)
	{
		show_loginbox();
		return;
	}
	if (got_downloading)
	{
		highlight_menu(1);
		hide_main();
		document.getElementById('download_box').style.display = '';
	} else {
		request_download();
	}
}

function do_nop()
{
	alert('В данной версии выполнение действий не реализованно');
}

function showHide(tnames)
{
	tname = tnames.split(';');
	for (i = 0; i < tname.length; i++)
	{
		var cobj = document.getElementById(tname[i]);
		if (cobj.style.display == '')
			cobj.style.display = 'none';
		else
			cobj.style.display = '';
	}
}