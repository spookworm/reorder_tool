function render() {
    body.innerHTML = "";
    const term =
        search.value
            .trim()
            .toLowerCase();
    filteredIndexes = [];
    fileData.forEach((item, index) => {
        const searchable =
            item.fields
                .map(value =>
                    String(value || "")
                )
                .join(" ")
                .toLowerCase();
        if (
            term === "" ||
            searchable.includes(term)
        ) {
            filteredIndexes.push(index);
        }
    });
    filteredIndexes.forEach(realIndex => {
        const item =
            fileData[realIndex];
        const tr =
            document.createElement("tr");
        tr.draggable =
            search.value.trim() === "";
        tr.dataset.id =
            item.id;
        let rowHtml = `
            <td class="reorder-cell">
                <div class="reorder-controls">
                    <button class="up-btn">
                        ↑
                    </button>
                    <input
                        type="number"
                        min="1"
                        max="${fileData.length}"
                        value="${item.reorder}"
                        class="move-input">
                    <button class="down-btn">
                        ↓
                    </button>
                    <button class="move-go">
                        →
                    </button>
                </div>
            </td>
            <td>
                ${item.number}
            </td>
        `;
        rowHtml += columnOrder.map(index => {
            const value =
                item.fields[index] || "";
            const heading =
                normalizeText(
                    importedHeadings[index]
                );
            let content;
            /*
                Special always-visible dropdown
            */
            if (
                index === toReadColumnIndex
            ){
                content = `
                <select class="to-read-editor">
                    <option value="0"
                    ${
                        String(value) === "0"
                        ? "selected"
                        : ""
                    }>
                        0
                    </option>
                    <option value="1"
                    ${
                        String(value) === "1"
                        ? "selected"
                        : ""
                    }>
                        1
                    </option>
                </select>
                `;
            }
            else {
                content = `
                <div class="cell-display">
                    ${escapeHtml(value)}
                </div>
                `;
            }
            const width =
                columnSizes[index] ||
                DEFAULT_COLUMN_WIDTH;
            return `
            <td
                data-column="${index}"
                style="
                    width:${width}px;
                    min-width:${width}px;
                    max-width:${width}px;
                "
            >
                ${content}
            </td>
            `;
        }).join("");
        rowHtml += `
            <td class="scd-column">
                ${escapeHtml(item.start_date)}
            </td>
            <td class="scd-column">
                ${escapeHtml(item.end_date)}
            </td>
            <td>
                ${escapeHtml(item.record_deleted_flag)}
            </td>
        `;
        tr.innerHTML =
            rowHtml;
        /*
            TAP TO EDIT
        */
        tr.querySelectorAll(".cell-display")
        .forEach(display => {
            display.onclick = () => {
                const td =
                    display.closest("td");
                const column =
                    Number(
                        td.dataset.column
                    );
                const heading =
                    normalizeText(
                        importedHeadings[column]
                    );
                const oldValue =
                    item.fields[column] || "";
                let editor;
                /*
                    Exclusive Shelf
                */
                if (
                    heading ===
                    "exclusive shelf"
                ){
                    editor =
                        document.createElement(
                            "select"
                        );
                    [
                        "to-read",
                        "ignore",
                        "read",
                        "currently-reading"
                    ]
                    .forEach(optionValue => {
                        const option =
                            document.createElement(
                                "option"
                            );
                        option.value =
                            optionValue;
                        option.textContent =
                            optionValue;
                        if(
                            optionValue === oldValue
                        ){
                            option.selected =
                                true;
                        }
                        editor.appendChild(
                            option
                        );
                    });
                }
                /*
                    My Rating
                */
                else if (
                    heading === "my rating"
                ){
                    editor =
                        document.createElement(
                            "input"
                        );
                    editor.type =
                        "number";
                    editor.min =
                        "-10";
                    editor.max =
                        "10";
                    editor.step =
                        "1";
                    editor.value =
                        oldValue;
                    editor.onchange =
                        () => {
                            let value =
                                Number(
                                    editor.value
                                );
                            if(
                                value < -10
                            ){
                                value = -10;
                            }
                            if(
                                value > 10
                            ){
                                value = 10;
                            }
                            if(
                                !Number.isInteger(value)
                            ){
                                value = 0;
                            }
                            editor.value =
                                value;
                        };
                }
                /*
                    Private Notes
                */
                else if (
                    heading ===
                    "private notes"
                ){
                    editor =
                        document.createElement(
                            "textarea"
                        );
                    editor.value =
                        oldValue;
                }
                /*
                    Prerequisite Book ID Goodreads
                */
				else if (
					heading ===
					"prerequisite - book id - goodreads"
				){
					const wrapper =
						document.createElement("div");
					wrapper.className =
						"prerequisite-wrapper";
					editor =
						document.createElement("input");
					editor.type =
						"text";
					editor.value =
						oldValue;
					editor.className =
						"prerequisite-editor";
					const suggestionBox =
						document.createElement("div");
					suggestionBox.className =
						"prerequisite-suggestions";
					wrapper.appendChild(editor);
					wrapper.appendChild(suggestionBox);
					td.innerHTML = "";
					td.appendChild(wrapper);
					editor.focus();
					editor.oninput = () => {
						suggestionBox.innerHTML = "";
						const searchText =
							editor.value
							.split("#")
							.pop()
							.trim()
							.toLowerCase();
						if(!searchText){
							return;
						}
						getBookTitleSuggestions()
						.filter(book =>
							book.title
							.toLowerCase()
							.includes(searchText)
						)
						.slice(0,10)
						.forEach(book => {
							const option =
								document.createElement("div");
							option.className =
								"prerequisite-option";
							option.textContent =
								`${book.title} (${book.id})`;
							option.onclick = function(e){
								e.preventDefault();
								e.stopPropagation();
								let parts =
									editor.value
									.split("#")
									.map(x => x.trim())
									.filter(Boolean);
								// remove the typed search text
								parts.pop();
								// add the selected Goodreads ID
								parts.push(
									String(book.id)
								);
								editor.value =
									parts.join("#") + "#";
								suggestionBox.innerHTML = "";
								setTimeout(() => {
									editor.focus();
								}, 0);
							};
							suggestionBox.appendChild(option);
						});
					};
					const savePrerequisite = () => {
						updateCell(
							item,
							column,
							editor.value
						);
						render();
					};
					editor.onblur = () => {
						setTimeout(() => {
							updateCell(
								item,
								column,
								editor.value
							);
							render();
						},200);
					};
					editor.onkeydown =
						event => {
							if(event.key === "Enter"){
								event.preventDefault();
								savePrerequisite();
							}
						};
					return;
				}
                /*
                    Normal text fields
                */
                else {
                    editor =
                        document.createElement(
                            "input"
                        );
                    editor.type =
                        "text";
                    editor.value =
                        oldValue;
                }
                editor.className =
                    "cell-editing";
                td.innerHTML =
                    "";
                td.appendChild(
                    editor
                );
                editor.focus();
                const saveEdit =
                () => {
                    let value =
                        editor.value;
                    if(
                        heading ===
                        "my rating"
                    ){
                        let number =
                            Number(value);
                        if(
                            !Number.isInteger(number)
                        ){
                            number = 0;
                        }
                        if(
                            number < -10
                        ){
                            number = -10;
                        }
                        if(
                            number > 10
                        ){
                            number = 10;
                        }
                        value =
                            number;
                    }
                    updateCell(
                        item,
                        column,
                        value
                    );
                    render();
                };
                editor.onchange =
                    saveEdit;
                editor.onblur =
                    saveEdit;
            };
        });
        body.appendChild(tr);
        /*
            TO READ DROPDOWN
        */
        const toReadEditor =
            tr.querySelector(
                ".to-read-editor"
            );
        if(toReadEditor){
            toReadEditor.onchange =
            event => {
                updateCell(
                    item,
                    toReadColumnIndex,
                    event.target.value
                );
            };
        }
        /*
            STOP BUTTON CLICKS
        */
        const moveInput =
            tr.querySelector(
                ".move-input"
            );
        if(moveInput){
            moveInput.onclick =
            event => {
                event.stopPropagation();
            };
        }
        /*
            MOVE UP
        */
        const up =
            tr.querySelector(
                ".up-btn"
            );
        if(up){
            up.onclick =
            event => {
                event.preventDefault();
                event.stopPropagation();
                move(
                    realIndex,
                    -1
                );
            };
        }
        /*
            MOVE DOWN
        */
        const down =
            tr.querySelector(
                ".down-btn"
            );
        if(down){
            down.onclick =
            event => {
                event.preventDefault();
                event.stopPropagation();
                move(
                    realIndex,
                    1
                );
            };
        }
        /*
            MOVE TO POSITION
        */
        const moveGo =
            tr.querySelector(
                ".move-go"
            );
        if(moveGo){
            moveGo.onclick =
            event => {
                event.preventDefault();
                event.stopPropagation();
                moveTo(
                    realIndex,
                    tr.querySelector(
                        ".move-input"
                    ).value
                );
            };
        }
        /*
            DRAG ROW REORDER
        */
        if(tr.draggable){
            attachRowDragHandlers(
                tr,
                item
            );
        }
    });
    info.textContent =
        `Loaded: ${
            originalFileName || "none"
        } | Rows: ${
            fileData.length
        }`;
}
function attachRowDragHandlers(
    tr,
    item
){
    tr.addEventListener(
        "dragstart",
        event => {
            snapshotState();
            event.dataTransfer.setData(
                "row",
                item.id
            );
            tr.classList.add(
                "dragging"
            );
        }
    );
    tr.addEventListener(
        "dragend",
        () => {
            tr.classList.remove(
                "dragging"
            );
        }
    );
    tr.addEventListener(
        "dragover",
        event => {
            event.preventDefault();
        }
    );
    tr.addEventListener(
        "drop",
        event => {
            event.preventDefault();
            const from =
                fileData.findIndex(
                    row =>
                        row.id ===
                        event.dataTransfer.getData(
                            "row"
                        )
                );
            const to =
                fileData.findIndex(
                    row =>
                        row.id ===
                        tr.dataset.id
                );
            if(
                from === -1 ||
                to === -1 ||
                from === to
            ){
                return;
            }
            const row =
                fileData.splice(
                    from,
                    1
                )[0];
            fileData.splice(
                to,
                0,
                row
            );
            updateReorderNumbers();
            fileData.forEach(
                (row,index)=>{
                    row.originalPosition=index;
                }
            );
            render();
        }
    );
}
function buildColumnGroup(){
    const group =
        document.getElementById(
            "columnGroup"
        );
    if(!group){
        return;
    }
    group.innerHTML = "";
    [
        140,
        40
    ].forEach(width => {
        const col =
            document.createElement("col");
        col.style.width =
            width + "px";
        group.appendChild(col);
    });
    columnOrder.forEach(index => {
        const col =
            document.createElement("col");
        col.dataset.column =
            index;
        col.style.width =
            (
                columnSizes[index] ||
                DEFAULT_COLUMN_WIDTH
            ) + "px";
        group.appendChild(col);
    });
    [
        180,
        180,
        180
    ].forEach(width => {
        const col =
            document.createElement("col");
        col.style.width =
            width + "px";
        group.appendChild(col);
    });
}
function renderHeaders(){
    const row =
        document.getElementById(
            "headers"
        );
    if(!row){
        console.error(
            "Header row not found"
        );
        return;
    }
    row.innerHTML = "";
    buildColumnGroup();
    /*
        Fixed columns
    */
    FIXED_COLUMNS.forEach(title => {
        const th =
            document.createElement("th");
        th.textContent =
            title;
        th.className =
            "fixed-column";
        row.appendChild(th);
    });
    /*
        Imported columns
    */
    columnOrder.forEach(columnIndex => {
        const th =
            document.createElement("th");
        th.className =
            "column-draggable";
        th.draggable = true;
        th.dataset.column =
            columnIndex;
        th.dataset.sortDirection =
            "1";
        th.innerHTML = `
            ${escapeHtml(
                importedHeadings[columnIndex]
            )}
            <span class="resize-handle"></span>
        `;
        row.appendChild(th);
        attachResizeHandlers(
            th,
            columnIndex
        );
        attachDragHandlers(
            th,
            columnIndex
        );
        let dragged = false;
        th.addEventListener(
            "dragstart",
            () => {
                dragged = true;
            }
        );
        th.addEventListener(
            "dragend",
            () => {
                setTimeout(
                    ()=>{
                        dragged = false;
                    },
                    50
                );
            }
        );
        th.addEventListener(
            "click",
            () => {
                if(dragged){
                    return;
                }
                const direction =
                    Number(
                        th.dataset.sortDirection
                    );
                sortByColumn(
                    columnIndex,
                    direction
                );
                th.dataset.sortDirection =
                    direction === 1
                    ? "-1"
                    : "1";
            }
        );
    });
    /*
        SCD columns
    */
    SCD_COLUMNS.forEach(title => {
        const th =
            document.createElement("th");
        th.textContent =
            title;
        th.className =
            "fixed-column";
        row.appendChild(th);
    });
}