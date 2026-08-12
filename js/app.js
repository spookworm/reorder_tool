let fileData = [];
let historyData = [];
let undoStack = [];

let importedHeadings = [];
let originalSnapshot = [];
let originalOrder = [];

let filteredIndexes = [];

let originalFileName = "";
let modifiedFileName = "";

let fileDelimiter = "\t";

let columnOrder = [];
let columnSizes = [];

let draggedColumn = null;
let resizingColumn = null;
let resizeStartX = 0;
let resizeStartWidth = 0;

let hasReorderColumn = false;
let hasNumberColumn = false;
let toReadColumnIndex = -1;
let editableTextColumns = [];

const body = document.getElementById("body");
const search = document.getElementById("search");
const save = document.getElementById("save");
const info = document.getElementById("info");

window.addEventListener("DOMContentLoaded", () => {
    document.getElementById("open").onclick = openFile;
    save.onclick = saveFile;
    search.oninput = render;
});

function snapshotState() {
	undoStack.push(JSON.stringify(fileData));

	if (undoStack.length > 50) {
		undoStack.shift();
	}
}

function undo() {
	const previous = undoStack.pop();

	if (!previous) {
		return;
	}

	fileData = JSON.parse(previous);
	render();
}

function markChanged(row) {
	row.lastModified = new Date().toISOString();
}

function getColumnName(index) {
	return importedHeadings[index] || "";
}