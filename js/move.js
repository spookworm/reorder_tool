function move(index, direction) {
    const target = index + direction;
    if (
        index < 0 ||
        index >= fileData.length ||
        target < 0 ||
        target >= fileData.length
    ) {
        return;
    }
    snapshotState();
    [
        fileData[index],
        fileData[target]
    ] = [
        fileData[target],
        fileData[index]
    ];
    updateReorderNumbers();
    render();
}
function moveTo(index, value) {
    const target = Number.parseInt(value, 10) - 1;
    if (
        Number.isNaN(target) ||
        index < 0 ||
        index >= fileData.length ||
        target < 0 ||
        target >= fileData.length ||
        target === index
    ) {
        return;
    }
    snapshotState();
    const row = fileData.splice(index, 1)[0];
    fileData.splice(target, 0, row);
    updateReorderNumbers();
    render();
}
function resetOrder() {
    snapshotState();
    const orderMap = new Map(
        originalOrder.map(
            (id, index) => [id, index]
        )
    );
    fileData.sort(
        (a, b) =>
            orderMap.get(a.id) -
            orderMap.get(b.id)
    );
    updateReorderNumbers();
    render();
}
function updateReorderNumbers() {
    fileData.forEach((row, index) => {
        row.number = index + 1;
        row.reorder = index + 1;
    });
}
function sortData(field, direction) {
    snapshotState();
    fileData.sort((a, b) => {
        let x = "";
        let y = "";
        if (field === "fields") {
            x = a.fields.join(" ");
            y = b.fields.join(" ");
        } else {
            x = a[field] || "";
            y = b[field] || "";
        }
        return compareValues(
            x,
            y,
            direction
        );
    });
    render();
}
function moveRowById(id, targetIndex) {
    const fromIndex = fileData.findIndex(
        row => row.id === id
    );
    if (
        fromIndex === -1 ||
        targetIndex < 0 ||
        targetIndex >= fileData.length ||
        fromIndex === targetIndex
    ) {
        return;
    }
    snapshotState();
    const row = fileData.splice(
        fromIndex,
        1
    )[0];
    fileData.splice(
        targetIndex,
        0,
        row
    );
    updateReorderNumbers();
    render();
}
function sortByColumn(columnIndex, direction = 1) {
    snapshotState();
    fileData.sort((a, b) => {
        const x = String(a.fields[columnIndex] || "");
        const y = String(b.fields[columnIndex] || "");
        return x.localeCompare(
            y,
            undefined,
            {
                numeric: true,
                sensitivity: "base"
            }
        ) * direction;
    });
    render();
}
function sortByRank(direction = 1) {
    snapshotState();
    fileData.sort((a, b) => {
        return (
            (Number(a.reorder) || 0) -
            (Number(b.reorder) || 0)
        ) * direction;
    });
    render();
}
function initialSCDSort() {
    if (!Array.isArray(fileData)) {
        return;
    }
    fileData.sort((a, b) => {
        const aDeleted =
            String(a.record_deleted_flag) === "1";
        const bDeleted =
            String(b.record_deleted_flag) === "1";
        // Active records first
        if (aDeleted !== bDeleted) {
            return aDeleted ? 1 : -1;
        }
        // Older start dates first
        const dateCompare =
            String(a.start_date || "")
                .localeCompare(
                    String(b.start_date || "")
                );
        if (dateCompare !== 0) {
            return dateCompare;
        }
        // Preserve original order
        return (
            (a.originalPosition || 0) -
            (b.originalPosition || 0)
        );
    });
    fileData.forEach((row, index) => {
        row.originalPosition = index;
        if (
            row.number === null ||
            row.number === undefined ||
            Number.isNaN(row.number)
        ) {
            row.number = index + 1;
        }
        if (
            row.reorder === null ||
            row.reorder === undefined ||
            Number.isNaN(row.reorder)
        ) {
            row.reorder = row.number;
        }
    });
}