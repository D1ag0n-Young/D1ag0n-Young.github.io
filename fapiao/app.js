pdfjsLib.GlobalWorkerOptions.workerSrc = 'fapiao/lib/pdf.worker.min.js';

var invoices = [];
var perPage = 2;
var seq = 0;
var pending = 0;
var dragState = null;

var firstBatchDone = false;
var batchDecision = null;

function isIE() {
	return !!window.ActiveXObject || 'ActiveXObject' in window;
}

$(function () {
	initFooter();
	if (isIE()) {
		$('#sys-alert').show();
		$('#sys-alert').animate({ 'height': '100%' }, 1000);
	}
	bindUI();
	bindPageDrop();
	applyPerPageUI();
	render();
});

function initFooter() {
	$('#footerYear').text(new Date().getFullYear());
}

function noteBatch() {
	if (batchDecision !== null) return;
	batchDecision = !firstBatchDone;
	firstBatchDone = true;
	setTimeout(function () { batchDecision = null; }, 0);
}

/* ---------------- 每页张数 ---------------- */

function setPerPage(n) {
	perPage = n;
	applyPerPageUI();
	renderPreview();
}

function applyPerPageUI() {
	$('#perPageSeg button').removeClass('active').filter('[data-n="' + perPage + '"]').addClass('active');
	$('#orient').html(perPage === 4
		? '@page { size: A4 landscape; margin:0 !important;} '
		: '@page { size: A4 portrait; margin:0 !important;} ');
}

/* ---------------- 发票渲染 ---------------- */

function beginRender() {
	pending++;
	updateParseHint();
}

function endRender() {
	pending = Math.max(0, pending - 1);
	updateParseHint();
}

function updateParseHint() {
	$('#parseHint').text(pending > 0 ? '解析中 ' + pending + ' …' : '');
}

function addInvoice(canvas, file) {
	invoices.push({
		id: ++seq,
		name: file.name,
		file: file,
		selected: !!file.__defaultSelected,
		src: canvas.toDataURL('image/jpeg', 0.92),
		thumb: makeThumb(canvas, 160)
	});
	render();
	endRender();
}

function makeThumb(canvas, w) {
	var h = Math.max(1, Math.round(canvas.height * w / canvas.width));
	var c = document.createElement('canvas');
	c.width = w;
	c.height = h;
	var ctx = c.getContext('2d');
	ctx.fillStyle = '#ffffff';
	ctx.fillRect(0, 0, w, h);
	ctx.drawImage(canvas, 0, 0, w, h);
	return c.toDataURL('image/jpeg', 0.7);
}

function readPdf(file) {
	beginRender();
	var reader = new FileReader();
	reader.readAsArrayBuffer(file);
	reader.onload = function (e) {
		var task = pdfjsLib.getDocument({ 'data': new Uint8Array(e.target.result), 'cMapUrl': 'sy/cmaps/', 'cMapPacked': true });
		task.promise.then(function (pdf) {
			pdf.getPage(1).then(function (page) {
				var viewport = page.getViewport({ 'scale': 3 });
				var canvas = document.createElement('canvas');
				canvas.width = Math.round(viewport.width);
				canvas.height = Math.round(viewport.height);
				var ctx = canvas.getContext('2d');
				ctx.fillStyle = '#ffffff';
				ctx.fillRect(0, 0, canvas.width, canvas.height);
				page.render({ 'canvasContext': ctx, 'viewport': viewport }).promise.then(function () {
					addInvoice(canvas, file);
				}, function () { endRender(); });
			}, function () { endRender(); });
		}, function () { endRender(); });
	};
	reader.onerror = function () { endRender(); };
}

function readImage(file) {
	beginRender();
	var url = URL.createObjectURL(file);
	var img = new Image();
	img.onload = function () {
		var scale = Math.min(1, 2400 / img.width);
		var canvas = document.createElement('canvas');
		canvas.width = Math.round(img.width * scale);
		canvas.height = Math.round(img.height * scale);
		var ctx = canvas.getContext('2d');
		ctx.fillStyle = '#ffffff';
		ctx.fillRect(0, 0, canvas.width, canvas.height);
		ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
		URL.revokeObjectURL(url);
		addInvoice(canvas, file);
	};
	img.onerror = function () {
		URL.revokeObjectURL(url);
		endRender();
	};
	img.src = url;
}

/* ---------------- 列表与预览 ---------------- */

function render() {
	renderList();
	renderPreview();
}

function renderList() {
	var list = $('#invoiceList').empty();
	invoices.forEach(function (inv) {
		var li = $('<li class="inv-item"></li>').attr('data-id', inv.id);
		if (!inv.selected) li.addClass('off');
		li.append($('<label class="inv-check"></label>').append($('<input type="checkbox">').prop('checked', inv.selected)));
		li.append($('<span class="inv-drag" title="拖动调整顺序">≡</span>'));
		li.append($('<img class="inv-thumb">').attr('src', inv.thumb).attr('alt', inv.name));
		li.append($('<span class="inv-name"></span>').text(shortName(inv.name)).attr('title', inv.name));
		li.append($('<button type="button" class="inv-del" title="删除">×</button>'));
		list.append(li);
	});
	if (!invoices.length) list.append('<li class="list-empty">还没有发票，请从右下角上传</li>');
	updateCounts();
}

function renderPreview() {
	var host = $('#sheets').empty();
	var empty = $('#previewEmpty');
	var list = invoices.filter(function (v) { return v.selected; });
	updateCounts();

	if (!invoices.length) {
		empty.removeClass('warn').addClass('show');
		$('#peTitle').text('把发票 PDF 拖到右下角上传窗口');
		$('#peSub').text('自动合并、批量打印，最多节省 75% 纸张');
		return;
	}
	if (!list.length) {
		empty.addClass('warn show');
		$('#peTitle').text('没有勾选任何发票');
		$('#peSub').text('在右侧列表中勾选要参与合并的发票');
		return;
	}
	empty.removeClass('show');

	var pages = Math.ceil(list.length / perPage);
	for (var p = 0; p < pages; p++) {
		var sheet = $('<div class="sheet"></div>').addClass('per-' + perPage);
		for (var k = 0; k < perPage; k++) {
			var slot = $('<div class="slot"></div>');
			var idx = p * perPage + k;
			if (idx < list.length) slot.append($('<img>').attr('src', list[idx].src));
			sheet.append(slot);
		}
		appendCuts(sheet);
		host.append(
			$('<div class="sheet-wrap"></div>')
				.append(sheet)
				.append($('<div class="sheet-label"></div>').text('第 ' + (p + 1) + ' 页 / 共 ' + pages + ' 页'))
		);
	}
}

function appendCuts(sheet) {
	if (perPage === 2) {
		sheet.append('<div class="cut cut-h" style="top:50%"></div>');
	} else if (perPage === 3) {
		sheet.append('<div class="cut cut-h" style="top:33.333%"></div>');
		sheet.append('<div class="cut cut-h" style="top:66.667%"></div>');
	} else if (perPage === 4) {
		sheet.append('<div class="cut cut-h" style="top:50%"></div>');
		sheet.append('<div class="cut cut-v" style="left:50%"></div>');
	}
}

function updateCounts() {
	var sel = 0;
	for (var i = 0; i < invoices.length; i++) if (invoices[i].selected) sel++;
	$('#invCount').text(sel + ' / ' + invoices.length);
	$('#selectAllBtn').text(invoices.length && sel === invoices.length ? '全不选' : '全选');
}

function findInvoice(id) {
	for (var i = 0; i < invoices.length; i++) if (invoices[i].id === id) return invoices[i];
	return null;
}

function shortName(name, max) {
	max = max || 30;
	if (!name || name.length <= max) return name;
	var dot = name.lastIndexOf('.');
	var ext = (dot > 0 && name.length - dot <= 8) ? name.slice(dot) : '';
	var base = ext ? name.slice(0, dot) : name;
	var keep = Math.max(6, max - ext.length - 1);
	return base.slice(0, keep) + '…' + ext;
}

/* ---------------- 交互 ---------------- */

function bindUI() {
	$('#perPageSeg').on('click', 'button', function () { setPerPage(+$(this).attr('data-n')); });
	$('#printBtn').on('click', function () { window.print(); });
	$('#clearBtn').on('click', clearAll);
	$('#uploadToggle').on('click', function () {
		var box = $(this).closest('.merge-upload');
		box.toggleClass('collapsed');
		$(this).text(box.hasClass('collapsed') ? '展开 ▸' : '收起 ▾');
	});

	$('#selectAllBtn').on('click', function () {
		var sel = 0;
		for (var i = 0; i < invoices.length; i++) if (invoices[i].selected) sel++;
		var makeAll = !(invoices.length && sel === invoices.length);
		invoices.forEach(function (v) { v.selected = makeAll; });
		render();
	});

	$('#removeCheckedBtn').on('click', function () {
		var removed = invoices.filter(function (v) { return v.selected; });
		dropFromQueue(removed);
		invoices = invoices.filter(function (v) { return !v.selected; });
		render();
	});

	$('#invoiceList')
		.on('change', 'input[type="checkbox"]', function () {
			var inv = findInvoice(+$(this).closest('.inv-item').attr('data-id'));
			if (!inv) return;
			inv.selected = this.checked;
			$(this).closest('.inv-item').toggleClass('off', !inv.selected);
			renderPreview();
		})
		.on('click', '.inv-del', function () {
			var id = +$(this).closest('.inv-item').attr('data-id');
			var inv = findInvoice(id);
			if (inv) dropFromQueue([inv]);
			invoices = invoices.filter(function (v) { return v.id !== id; });
			render();
		})
		.on('contextmenu', '.inv-item', function (e) { e.preventDefault(); });

	bindDrag();
}

function dropFromQueue(list) {
	var dz = Dropzone.forElement('#pdfdz');
	if (!dz) return;
	list.forEach(function (inv) {
		try { if (inv.file) dz.removeFile(inv.file); } catch (e) { }
	});
}

function clearAll() {
	if (invoices.length && !window.confirm('确定清空全部发票？')) return;
	invoices = [];
	firstBatchDone = false;
	batchDecision = null;
	var dz = Dropzone.forElement('#pdfdz');
	if (dz) dz.removeAllFiles(true);
	render();
}

/* ---------------- 拖动排序 ---------------- */

function bindDrag() {
	var listEl = document.getElementById('invoiceList');
	listEl.addEventListener('pointerdown', function (e) {
		if (dragState) return;
		if (e.button != null && e.button !== 0) return;
		var item = e.target.closest ? e.target.closest('.inv-item') : null;
		if (!item) return;
		if (e.target.closest('.inv-check') || e.target.closest('.inv-del')) return;
		var onHandle = !!e.target.closest('.inv-drag');
		var rect = item.getBoundingClientRect();
		dragState = {
			item: item,
			id: +item.getAttribute('data-id'),
			startX: e.clientX,
			startY: e.clientY,
			rect: rect,
			active: false,
			needMove: false,
			ghost: null,
			placeholder: null,
			timer: null
		};
		if (onHandle) {
			dragState.timer = setTimeout(activateDrag, 0);
		} else if (e.pointerType === 'touch') {
			dragState.timer = setTimeout(activateDrag, 260);
		} else {
			dragState.needMove = true;
		}
		document.addEventListener('pointermove', onDragMove);
		document.addEventListener('pointerup', onDragEnd);
		document.addEventListener('pointercancel', onDragEnd);
		document.addEventListener('touchmove', onTouchMove, { passive: false });
	});
}

function activateDrag() {
	if (!dragState || dragState.active) return;
	var d = dragState;
	d.active = true;
	d.item.classList.add('drag-src');

	var ph = document.createElement('li');
	ph.className = 'drop-placeholder';
	ph.style.height = d.rect.height + 'px';
	d.item.parentNode.insertBefore(ph, d.item);
	d.item.style.display = 'none';
	d.placeholder = ph;

	var ghost = d.item.cloneNode(true);
	ghost.classList.remove('drag-src');
	ghost.classList.add('drag-ghost');
	ghost.style.display = '';
	ghost.style.width = d.rect.width + 'px';
	ghost.style.height = d.rect.height + 'px';
	ghost.style.left = d.rect.left + 'px';
	ghost.style.top = d.rect.top + 'px';
	document.body.appendChild(ghost);
	d.ghost = ghost;

	if (navigator.vibrate) { try { navigator.vibrate(12); } catch (e) { } }
}

function onTouchMove(e) {
	if (dragState && dragState.active) e.preventDefault();
}

function onDragMove(e) {
	if (!dragState) return;
	var d = dragState;
	if (!d.active) {
		var dx = Math.abs(e.clientX - d.startX);
		var dy = Math.abs(e.clientY - d.startY);
		if (d.needMove) {
			if (dx > 5 || dy > 5) activateDrag();
		} else if (dx > 10 || dy > 10) {
			cancelPress();
			return;
		}
		if (!d.active) return;
	}
	if (!document.contains(d.item)) { cancelPress(); return; }
	e.preventDefault();
	d.ghost.style.transform = 'translate(' + (e.clientX - d.startX) + 'px,' + (e.clientY - d.startY) + 'px)';

	var listEl = document.getElementById('invoiceList');
	var items = Array.prototype.slice.call(listEl.querySelectorAll('.inv-item'))
		.filter(function (el) { return el !== d.item; });
	var idx = items.length;
	for (var i = 0; i < items.length; i++) {
		var r = items[i].getBoundingClientRect();
		if (e.clientY < r.top + r.height / 2) { idx = i; break; }
	}
	if (idx >= items.length) listEl.appendChild(d.placeholder);
	else listEl.insertBefore(d.placeholder, items[idx]);

	var scroller = document.getElementById('listScroll');
	var sr = scroller.getBoundingClientRect();
	if (e.clientY < sr.top + 48) scroller.scrollTop -= 12;
	else if (e.clientY > sr.bottom - 48) scroller.scrollTop += 12;
}

function onDragEnd() {
	if (!dragState) return;
	if (!dragState.active) { cancelPress(); return; }
	finishDrag();
}

function cancelPress() {
	if (!dragState) return;
	if (dragState.timer) clearTimeout(dragState.timer);
	removeDragListeners();
	dragState = null;
}

function finishDrag() {
	var d = dragState;
	var ph = d.placeholder;
	var listEl = document.getElementById('invoiceList');
	var children = Array.prototype.slice.call(listEl.children);
	var to = 0;
	for (var i = 0; i < children.length; i++) {
		if (children[i] === ph) break;
		if (children[i].classList.contains('inv-item') && children[i] !== d.item) to++;
	}
	var from = -1;
	for (var j = 0; j < invoices.length; j++) if (invoices[j].id === d.id) { from = j; break; }

	d.item.style.display = '';
	d.item.classList.remove('drag-src');
	if (ph && ph.parentNode) ph.parentNode.removeChild(ph);
	if (d.ghost && d.ghost.parentNode) d.ghost.parentNode.removeChild(d.ghost);
	removeDragListeners();
	dragState = null;

	if (from >= 0 && from !== to) {
		var moved = invoices.splice(from, 1)[0];
		invoices.splice(to, 0, moved);
	}
	render();
}

function removeDragListeners() {
	document.removeEventListener('pointermove', onDragMove);
	document.removeEventListener('pointerup', onDragEnd);
	document.removeEventListener('pointercancel', onDragEnd);
	document.removeEventListener('touchmove', onTouchMove);
}

/* ---------------- 整页拖拽上传（上传窗口之外也可直接拖入） ---------------- */

function bindPageDrop() {
	var overlay = $('#dropOverlay');
	var depth = 0;

	function dt(e) { return e.originalEvent && e.originalEvent.dataTransfer; }
	function hasFiles(e) {
		var d = dt(e);
		if (!d || !d.types) return false;
		for (var i = 0; i < d.types.length; i++) if (d.types[i] === 'Files') return true;
		return false;
	}
	function hide() { depth = 0; overlay.removeClass('show'); }

	$(document)
		.on('dragenter', function (e) {
			if (!hasFiles(e)) return;
			e.preventDefault();
			depth++;
			overlay.addClass('show');
		})
		.on('dragover', function (e) {
			if (!hasFiles(e)) return;
			e.preventDefault();
			var d = dt(e);
			if (d) d.dropEffect = 'copy';
		})
		.on('dragleave', function (e) {
			if (!hasFiles(e)) return;
			if (!e.relatedTarget) { hide(); return; }
			depth = Math.max(0, depth - 1);
			if (!depth) overlay.removeClass('show');
		})
		.on('drop', function (e) {
			var files = dt(e) && dt(e).files;
			var onDropzone = $(e.target).closest('#pdfdz').length > 0;
			hide();
			if (!hasFiles(e)) return;
			e.preventDefault();
			if (onDropzone) return;
			var dz = Dropzone.forElement('#pdfdz');
			if (!dz || !files) return;
			for (var i = 0; i < files.length; i++) dz.addFile(files[i]);
		});
}

/* ---------------- 上传（逻辑保持不变） ---------------- */

Dropzone.options.pdfdz = {
	url: 'upload.jsp',
	method: 'post',
	paramName: 'file',
	maxFiles: 36,
	maxFilesize: 1,
	acceptedFiles: '.pdf, .png, .jpg, .jpeg',
	parallelUploads: 1,
	previewsContainer: false,
	dictMaxFilesExceeded: '您最多只能上传36个文件！',
	dictResponseError: '文件上传失败!',
	dictInvalidFileType: '上传的文件类型仅支持PDF',
	dictFallbackMessage: '浏览器不受支持',
	dictFileTooBig: '文件过大上传文件最大支持1M.',
	addedfiles: function () {
		noteBatch();
	},
	addedfile: function (file) {
		noteBatch();
		file.__defaultSelected = batchDecision;
	},
	error: function (file) {
		if (file.type === 'image/png' || file.type === 'image/jpeg') readImage(file);
		else readPdf(file);
	},
	success: function (file) {
		if (file.type === 'image/png' || file.type === 'image/jpeg') readImage(file);
		else readPdf(file);
	}
};
