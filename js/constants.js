const ACTIVE_END_DATE = "9998-12-31T23:59:59Z";

const EDITORS = {
    "private notes": "text",
    "prerequisite - book id - goodreads": "text"
};

const DESCRIPTION_COLUMN = "Description";
const DOI_COLUMN = "DOI";

const FIXED_COLUMNS = Object.freeze([
    "Reorder",
    "#"
]);

const SCD_COLUMNS = Object.freeze([
    "Start Date",
    "End Date",
    "Record Deleted Flag"
]);

const AUTO_FIT_COLUMNS = Object.freeze([
    "title",
    "author"
]);

const EXCLUSIVE_SHELF_VALUES = [
    "to-read",
    "ignore",
    "read",
    "currently-reading"
];

const MY_RATING_MIN = -10;
const MY_RATING_MAX = 10;

const MIN_COLUMN_WIDTH = 50;
const DEFAULT_COLUMN_WIDTH = 100;
const MAX_COLUMN_WIDTH = 400;