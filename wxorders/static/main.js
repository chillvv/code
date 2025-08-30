const dropArea = document.getElementById('drop-area');
const fileInput = document.getElementById('file-input');
const lunchStart = document.getElementById('lunchStart');
const dinnerStart = document.getElementById('dinnerStart');
const targetSel = document.getElementById('target');
const testMode = document.getElementById('testMode');
const previewLunch = document.getElementById('previewLunch');
const previewDinner = document.getElementById('previewDinner');

const btnPreview = document.getElementById('btnPreview');
const btnSendTest = document.getElementById('btnSendTest');
const btnSend = document.getElementById('btnSend');
const btnStop = document.getElementById('btnStop');

let currentFile = null;

function setHighlight(on) {
	if (on) dropArea.classList.add('highlight');
	else dropArea.classList.remove('highlight');
}

dropArea.addEventListener('dragover', (e) => {
	e.preventDefault();
	setHighlight(true);
});

dropArea.addEventListener('dragleave', () => setHighlight(false));

dropArea.addEventListener('drop', (e) => {
	e.preventDefault();
	setHighlight(false);
	const files = e.dataTransfer.files;
	if (files && files.length) {
		currentFile = files[0];
		fileInput.files = files;
	}
});

fileInput.addEventListener('change', () => {
	currentFile = fileInput.files && fileInput.files[0] ? fileInput.files[0] : null;
});

async function preview() {
	if (!currentFile) {
		alert('请先选择或拖入 Excel 文件');
		return;
	}
	const fd = new FormData();
	fd.append('file', currentFile);
	fd.append('lunchStart', String(lunchStart.value || 1));
	fd.append('dinnerStart', String(dinnerStart.value || 1));
	const res = await fetch('/api/preview', { method: 'POST', body: fd });
	if (!res.ok) {
		alert('预览失败');
		return;
	}
	const data = await res.json();
	previewLunch.textContent = data.lunch || '';
	previewDinner.textContent = data.dinner || '';
}

async function send(test) {
	const payload = {
		target: targetSel.value,
		textLunch: previewLunch.textContent || '',
		textDinner: previewDinner.textContent || '',
		test: test === true ? true : !(!testMode.checked)
	};
	const res = await fetch('/api/send', {
		method: 'POST',
		headers: { 'Content-Type': 'application/json' },
		body: JSON.stringify(payload)
	});
	if (!res.ok) {
		alert('发送启动失败');
		return;
	}
	alert('已开始发送，按 ESC 或点击停止可中止');
}

async function stopSending() {
	await fetch('/api/stop', { method: 'POST' });
}

btnPreview.addEventListener('click', preview);
btnSendTest.addEventListener('click', () => send(true));
btnSend.addEventListener('click', () => send(false));
btnStop.addEventListener('click', stopSending);

document.addEventListener('keydown', (e) => {
	if (e.key === 'Escape') {
		stopSending();
	}
});