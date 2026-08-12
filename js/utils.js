function escapeHtml(text) {
    if (text == null) {
        return "";
    }
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
function updateField(row, property, value) {
    if (!row) {
        return;
    }
    row[property] = value;
    row.lastModified =
        new Date()
            .toISOString();
}
function updateCell(row, columnIndex, value) {
    if (
        !row ||
        !Array.isArray(row.fields)
    ) {
        return;
    }
    row.fields[columnIndex] = value;
    row.lastModified =
        new Date()
            .toISOString();
}
function cloneData(data) {
    return JSON.parse(
        JSON.stringify(data)
    );
}
function createId() {
    return (
        Date.now().toString(36) +
        Math.random()
            .toString(36)
            .substring(2)
    );
}
function normalizeText(value) {
    return String(value || "")
        .trim()
        .toLowerCase();
}
function compareValues(
    a,
    b,
    direction = 1
) {
    return String(a || "")
        .localeCompare(
            String(b || ""),
            undefined,
            {
                numeric: true,
                sensitivity: "base"
            }
        ) * direction;
}
function snapshotState() {
    if (
        typeof historyStack === "undefined"
    ) {
        return;
    }
    historyStack.push(
        cloneData(fileData)
    );
    if (
        historyStack.length > 50
    ) {
        historyStack.shift();
    }
}
function validatePrerequisiteBookIds() {
    const bookIdColumn =
        importedHeadings.findIndex(
            h =>
                normalizeText(h) ===
                "book id - goodreads"
        );
    const prerequisiteColumn =
        importedHeadings.findIndex(
            h =>
                normalizeText(h) ===
                "prerequisite - book id - goodreads"
        );
    if (
        bookIdColumn === -1 ||
        prerequisiteColumn === -1
    ) {
        console.log("Validation columns not found");
        return true;
    }
    const validBookIds = new Set();
    fileData.forEach(row => {
        const value =
            String(
                row.fields[bookIdColumn] ?? ""
            )
            .trim();
        if (value !== "") {
            validBookIds.add(value);
        }
    });
    console.log(
        "Valid Book IDs:",
        Array.from(validBookIds)
    );
    const invalid = [];
    fileData.forEach((row, index) => {
        const prerequisite =
            String(
                row.fields[prerequisiteColumn] ?? ""
            )
            .trim();
        if (
            prerequisite !== "" &&
            !validBookIds.has(prerequisite)
        ) {
            invalid.push(
                `Row ${index + 1}: ${prerequisite}`
            );
        }
    });
    if (invalid.length > 0) {
        alert(
            "Invalid Prerequisite - Book Id - Goodreads values:\n\n" +
            invalid.join("\n")
        );
        return false;
    }
    return true;
}
function getBookIdSuggestions() {
    const bookIdColumn =
        importedHeadings.findIndex(
            heading =>
                normalizeText(heading) ===
                "book id - goodreads"
        );
    if (bookIdColumn === -1) {
        return [];
    }
    return [
        ...new Set(
            fileData
                .map(row =>
                    String(
                        row.fields[bookIdColumn] || ""
                    ).trim()
                )
                .filter(value => value !== "")
        )
    ];
}
function getBookTitleSuggestions() {
    const titleColumn =
        importedHeadings.findIndex(
            heading =>
                normalizeText(heading) === "title"
        );
    const bookIdColumn =
        importedHeadings.findIndex(
            heading =>
                normalizeText(heading) ===
                "book id - goodreads"
        );
    if (
        titleColumn === -1 ||
        bookIdColumn === -1
    ) {
        return [];
    }
    return fileData
        .map(row => ({
            title: String(
                row.fields[titleColumn] || ""
            ).trim(),
            id: String(
                row.fields[bookIdColumn] || ""
            ).trim()
        }))
        .filter(item =>
            item.title !== "" &&
            item.id !== ""
        );
}