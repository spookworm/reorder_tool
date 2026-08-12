function openFile() {
    const input =
        document.createElement("input");
    input.type = "file";
    input.accept =
        ".txt,.csv,.log,.md,.tsv";
    input.onchange = event => {
        const file =
            event.target.files[0];
        if (!file) {
            return;
        }
        originalFileName =
            file.name;
        const reader =
            new FileReader();
        reader.onload = () => {
            try {
                console.log(
                    "Importing:",
                    originalFileName
                );
                /*
                    Clear previous state
                    before loading new file
                */
                fileData = [];
                historyData = [];
                originalSnapshot = [];
                originalOrder = [];
                columnOrder = [];
                columnSizes = [];
                parse(
                    reader.result
                );
                console.log(
                    "Parsed rows:",
                    fileData.length
                );
                console.log(
                    "Columns:",
                    importedHeadings
                );
                if (
                    fileData.length === 0
                ) {
                    throw new Error(
                        "No data rows found"
                    );
                }
                /*
                    Validate ISBN / DOI /
                    Goodreads ID rule
                */
                if (
                    typeof validateBookIdentifiers ===
                    "function"
                ) {
                    if (
                        !validateBookIdentifiers()
                    ) {
                        return;
                    }
                }
                renderHeaders();
                modifiedFileName =
                    originalFileName.replace(
                        /\.[^/.]+$/,
                        ""
                    ) +
                    "_reordered_" +
                    new Date()
                        .toISOString()
                        .replace(
                            /[:.]/g,
                            "-"
                        ) +
                    ".txt";
                if (save) {
                    save.disabled =
                        false;
                }
                render();
                console.log(
                    "Import complete"
                );
            }
            catch(error) {
                console.error(
                    "IMPORT ERROR:",
                    error
                );
                alert(
                    "Import failed:\n\n" +
                    error.message
                );
            }
        };
        reader.onerror = () => {
            alert(
                "Could not read file"
            );
        };
        reader.readAsText(file);
    };
    input.click();
}
function validateBookIdentifiers() {
    const isbnIndex =
        importedHeadings.indexOf("ISBN");
    const doiIndex =
        importedHeadings.indexOf("DOI");
    const bookIdIndex =
        importedHeadings.indexOf(
            "Book Id - Goodreads"
        );
    const invalidRows = [];
    fileData.forEach((row, index) => {
        const isbn =
            String(
                row.fields[isbnIndex] || ""
            ).trim();
        const doi =
            String(
                row.fields[doiIndex] || ""
            ).trim();
        const bookId =
            String(
                row.fields[bookIdIndex] || ""
            ).trim();
        if (
            isbn === "" &&
            doi === "" &&
            bookId === ""
        ) {
            invalidRows.push(
                index + 1
            );
        }
    });
    if (invalidRows.length > 0) {
        alert(
            "The following rows have no ISBN, DOI, or Book Id - Goodreads:\n\n" +
            invalidRows.join(", ")
        );
        return false;
    }
    return true;
}
function parse(text) {
    importedHeadings = [];
    fileData = [];
    historyData = [];
    originalOrder = [];
    originalSnapshot = [];
    columnOrder = [];
    columnSizes = [];
    hasNumberColumn = false;
    toReadColumnIndex = -1;
    editableTextColumns = [];
    const lines =
        text
            .split(/\r?\n/)
            .filter(line => line.trim());
    if (lines.length === 0) {
        alert("File is empty");
        return;
    }
    fileDelimiter =
        lines[0].includes("\t")
            ? "\t"
            : lines[0].includes("|")
                ? "|"
                : ",";
    let headings =
        lines[0]
            .split(fileDelimiter)
            .map(value => value.trim());
    const isHeader =
        headings.some(
            value =>
                /[A-Za-z]/.test(value)
        );
    if (!isHeader) {
        importedHeadings =
            headings.map(
                (_, index) =>
                    "Column " + (index + 1)
            );
    }
    else {
        importedHeadings =
            [...headings];
    }
    lines.shift();
    /*
        Remove # column
    */
    const numberIndex =
        importedHeadings.findIndex(
            heading =>
                normalizeText(heading) === "#"
        );
    if (numberIndex !== -1) {
        hasNumberColumn = true;
        importedHeadings.splice(
            numberIndex,
            1
        );
    }
    /*
        Locate SCD columns
    */
    const startIndex =
        importedHeadings.findIndex(
            heading =>
                normalizeText(heading) ===
                "start date"
        );
    const endIndex =
        importedHeadings.findIndex(
            heading =>
                normalizeText(heading) ===
                "end date"
        );
    const deletedIndex =
        importedHeadings.findIndex(
            heading =>
                normalizeText(heading) ===
                "record deleted flag"
        );
    const hasSCDColumns =
        startIndex !== -1 &&
        endIndex !== -1 &&
        deletedIndex !== -1;
    /*
        Remove SCD headings from
        displayed columns
    */
    importedHeadings =
        importedHeadings.filter(
            heading => {
                const name =
                    normalizeText(heading);
                return (
                    name !== "start date" &&
                    name !== "end date" &&
                    name !== "record deleted flag"
                );
            }
        );
	// Every imported column is editable
	editableTextColumns =
		importedHeadings.map(
			(_, index) => index
		);
    toReadColumnIndex =
        importedHeadings.findIndex(
            heading =>
                normalizeText(heading) ===
                "to_read"
        );
    /*
        Import rows
    */
    lines.forEach(line => {
        let parts =
            line.split(fileDelimiter);
        let number =
            null;
        if (hasNumberColumn) {
            number =
                Number(
                    parts.shift()
                );
        }
        let startDate =
            "";
        let endDate =
            "";
        let deletedFlag =
            "0";
        if (hasSCDColumns) {
            startDate =
                parts[
                    parts.length - 3
                ];
            endDate =
                parts[
                    parts.length - 2
                ];
            deletedFlag =
                parts[
                    parts.length - 1
                ];
        }
        const hasValidSCD =
            hasSCDColumns &&
            /^\d{4}-\d{2}-\d{2}/.test(startDate) &&
            /^\d{4}-\d{2}-\d{2}/.test(endDate) &&
            (
                deletedFlag === "0" ||
                deletedFlag === "1"
            );
        let fields =
            hasValidSCD
                ? parts.slice(
                    0,
                    parts.length - 3
                )
                : parts;
        const now =
            new Date()
                .toISOString();
        fileData.push({
            id:
                createId(),
            fields,
            number,
            reorder:
                number,
            start_date:
                hasValidSCD
                    ? startDate
                    : now,
            end_date:
                hasValidSCD
                    ? endDate
                    : ACTIVE_END_DATE,
            record_deleted_flag:
                hasValidSCD
                    ? deletedFlag
                    : "0",
            lastModified:
                now
        });
    });
    /*
        Fix row numbering
    */
    fileData.forEach(
        (row,index)=>{
            row.originalPosition =
                index;
            if (
                row.number === null ||
                Number.isNaN(row.number)
            ) {
                row.number =
                    index + 1;
            }
            row.reorder =
                row.number;
        }
    );
    /*
        Insert DOI after ISBN
    */
    if (
        !importedHeadings.includes(
            "DOI"
        )
    ) {
        const isbnIndex =
            importedHeadings.indexOf(
                "ISBN"
            );
        const insertAt =
            isbnIndex === -1
                ? importedHeadings.length
                : isbnIndex + 1;
        importedHeadings.splice(
            insertAt,
            0,
            "DOI"
        );
        fileData.forEach(row=>{
            row.fields.splice(
                insertAt,
                0,
                ""
            );
        });
    }
    /*
        Add Description last
    */
    if (
        !importedHeadings.includes(
            "Description"
        )
    ) {
        importedHeadings.push(
            "Description"
        );
        fileData.forEach(row=>{
            row.fields.push("");
        });
    }
    /*
        Column order:
        Description always last
    */
    columnOrder =
        importedHeadings.map(
            (_,index)=>index
        );
    const descriptionIndex =
        importedHeadings.indexOf(
            "Description"
        );
    if (descriptionIndex !== -1) {
        columnOrder =
            columnOrder.filter(
                index =>
                    index !== descriptionIndex
            );
        columnOrder.push(
            descriptionIndex
        );
    }
    columnSizes =
        importedHeadings.map(
            ()=>null
        );
    autoFitImportantColumns();
    originalOrder =
        fileData.map(
            row =>
                row.id
        );
    /*
        Snapshot AFTER all added columns
    */
    originalSnapshot =
        cloneData(
            fileData
        );
}
function validateExclusiveShelf() {
	const column =
		importedHeadings.findIndex(
			heading =>
				normalizeText(heading) ===
				"exclusive shelf"
		);
	if (column === -1) {
		return true;
	}
	const invalid = [];
	fileData.forEach((row,index)=>{
		const value =
			String(
				row.fields[column] || ""
			).trim();
		if (
			value !== "" &&
			!EXCLUSIVE_SHELF_VALUES.includes(value)
		) {
			invalid.push(
				index + 1
			);
		}
	});
	if (invalid.length) {
		alert(
			"Invalid Exclusive Shelf values in rows:\n\n" +
			invalid.join(", ") +
			"\n\nAllowed values:\n" +
			EXCLUSIVE_SHELF_VALUES.join(", ")
		);
		return false;
	}
	return true;
}
function saveFile() {
	if (!validateMyRating()) {
		return;
	}
    if (!validatePrerequisiteBookIds()) {
        return;
    }
    const saveTime =
        new Date().toISOString();
    let activeOutput = [];
    let inactiveOutput = [];
    historyData.forEach(row => {
        inactiveOutput.push(
            [
                row.fields.join(fileDelimiter),
                row.start_date,
                row.end_date,
                row.record_deleted_flag
            ].join(fileDelimiter)
        );
    });
    fileData.forEach(item => {
        const old =
            originalSnapshot.find(
                row =>
                    row.id === item.id
            );
        const changed =
            !old ||
            JSON.stringify(old.fields) !==
                JSON.stringify(item.fields) ||
            old.number !== item.number ||
            old.reorder !== item.reorder;
        if (
            changed &&
            old
        ) {
            inactiveOutput.push(
                [
                    old.fields.join(fileDelimiter),
                    old.start_date,
                    saveTime,
                    "1"
                ].join(fileDelimiter)
            );
        }
        const exportedFields =
            columnOrder.map(
                index =>
                    item.fields[index] ?? ""
            );
        activeOutput.push(
            [
                exportedFields.join(fileDelimiter),
                changed
                    ? saveTime
                    : item.start_date,
                ACTIVE_END_DATE,
                "0"
            ].join(fileDelimiter)
        );
    });
    const outputHeadings =
        columnOrder.map(
            index =>
                importedHeadings[index]
        );
    outputHeadings.push(
        "Start Date",
        "End Date",
        "Record Deleted Flag"
    );
    const output = [
        outputHeadings.join(fileDelimiter),
        ...activeOutput,
        ...inactiveOutput
    ];
    const blob =
        new Blob(
            [
                output.join("\n")
            ],
            {
                type: "text/plain"
            }
        );
    const url =
        URL.createObjectURL(blob);
    const a =
        document.createElement("a");
    a.href = url;
    a.download = modifiedFileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    alert(
        "File saved:\n" +
        modifiedFileName
    );
    originalSnapshot =
        JSON.parse(
            JSON.stringify(fileData)
        );
    fileData.forEach(
        (row, index) => {
            row.originalPosition = index;
        }
    );
}
function validateMyRating() {
    const column =
        importedHeadings.findIndex(
            heading =>
                normalizeText(heading) ===
                "my rating"
        );
    if (column === -1) {
        return true;
    }
    const invalidRows = [];
    fileData.forEach((row,index)=>{
        const value =
            String(
                row.fields[column] || ""
            ).trim();
        if (value === "") {
            return;
        }
        const number =
            Number(value);
        if (
            !Number.isInteger(number) ||
            number < MY_RATING_MIN ||
            number > MY_RATING_MAX
        ) {
            invalidRows.push(
                index + 1
            );
        }
    });
    if (invalidRows.length > 0) {
        alert(
            "Invalid My Rating values in rows:\n\n" +
            invalidRows.join(", ") +
            "\n\nMy Rating must be an integer between -10 and 10."
        );
        return false;
    }
    return true;
}