/**
 * Compare two dates.
 * @param a First date
 * @param b second date
 */
function compareDates(a, b){
    if(a > b)
        return 1;
    else if (a < b)
        return -1;
    else
        return 0;
}
