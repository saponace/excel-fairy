/**
 * Compare two dates.
 * @param a First date
 * @param b second date
 */
function compareDates(a, b){
    if(a > b)
        return 1;
    else
        return -1;
}


/**
 * Convert a potentially string date to Date object (if the date is a string, it has to be formatted as DD/MM/YYYY or D/MM/YYYY)
 * @param d date (either a string or a Date object
 * @return Date object
 */
function castStringToDateIfString(d){
    if(typeof d === 'string'){
        var values = d.split('/');
        return new Date(values[2], values[1]-1, values[0]);
    }
    else if (d instanceof Date)
        return d;
    else
        return null;
}