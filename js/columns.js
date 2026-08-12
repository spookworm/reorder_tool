function autoFitImportantColumns() {
    importedHeadings.forEach((heading, index) => {
        const name =
            normalizeText(heading);
        if (
            !AUTO_FIT_COLUMNS.includes(name)
        ) {
            return;
        }
        let maxLength =
            heading.length;
        fileData.forEach(row => {
            const value =
                row.fields[index] || "";
            if (
                String(value).length >
                maxLength
            ) {
                maxLength =
                    String(value).length;
            }
        });
        columnSizes[index] =
            Math.min(
                Math.max(
                    maxLength * 9 + 20,
                    DEFAULT_COLUMN_WIDTH
                ),
                MAX_COLUMN_WIDTH
            );
    });
}
function attachResizeHandlers(
    th,
    columnIndex
) {
    const handle =
        th.querySelector(
            ".resize-handle"
        );
    if (!handle) {
        return;
    }
    handle.addEventListener(
        "pointerdown",
        event => {
            event.preventDefault();
            event.stopPropagation();
            resizingColumn =
                columnIndex;
            resizeStartX =
                event.clientX;
            resizeStartWidth =
                columnSizes[columnIndex] ||
                th.offsetWidth;
            handle.setPointerCapture(
                event.pointerId
            );
            document.body.classList.add(
                "resizing"
            );
            function resize(moveEvent) {
                if (
                    resizingColumn === null
                ) {
                    return;
                }
                const width =
                    Math.max(
                        MIN_COLUMN_WIDTH,
                        resizeStartWidth +
                        (
                            moveEvent.clientX -
                            resizeStartX
                        )
                    );
                columnSizes[columnIndex] =
                    width;
                updateColumnWidth(
                    columnIndex,
                    width
                );
            }
            function stopResize() {
                resizingColumn =
                    null;
                document.body.classList.remove(
                    "resizing"
                );
                document.removeEventListener(
                    "pointermove",
                    resize
                );
                document.removeEventListener(
                    "pointerup",
                    stopResize
                );
                document.removeEventListener(
                    "pointercancel",
                    stopResize
                );
            }
            document.addEventListener(
                "pointermove",
                resize
            );
            document.addEventListener(
                "pointerup",
                stopResize
            );
            document.addEventListener(
                "pointercancel",
                stopResize
            );
        }
    );
}
function attachDragHandlers(
    th,
    columnIndex
) {
    th.addEventListener(
        "dragstart",
        event => {
            draggedColumn =
                columnIndex;
            event.dataTransfer.effectAllowed =
                "move";
            event.dataTransfer.setData(
                "column",
                String(columnIndex)
            );
            th.classList.add(
                "dragging"
            );
        }
    );
    th.addEventListener(
        "dragend",
        () => {
            draggedColumn =
                null;
            th.classList.remove(
                "dragging"
            );
        }
    );
    th.addEventListener(
        "dragover",
        event => {
            event.preventDefault();
            th.classList.add(
                "over"
            );
        }
    );
    th.addEventListener(
        "dragleave",
        () => {
            th.classList.remove(
                "over"
            );
        }
    );
    th.addEventListener(
        "drop",
        event => {
            event.preventDefault();
            th.classList.remove(
                "over"
            );
            const from =
                Number(
                    event.dataTransfer.getData(
                        "column"
                    )
                );
            const fromIndex =
                columnOrder.indexOf(
                    from
                );
            const targetIndex =
                columnOrder.indexOf(
                    columnIndex
                );
            if (
                fromIndex === -1 ||
                targetIndex === -1 ||
                fromIndex === targetIndex
            ) {
                return;
            }
            snapshotState();
            columnOrder.splice(
                fromIndex,
                1
            );
            let insertAt =
                targetIndex;
            if (
                fromIndex < targetIndex
            ) {
                insertAt--;
            }
            columnOrder.splice(
                insertAt,
                0,
                from
            );
            renderHeaders();
            render();
        }
    );
}