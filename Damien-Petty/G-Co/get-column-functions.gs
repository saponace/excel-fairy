function getColumnA(inputRow, inputSheetCode) {
    return inputRow[letterToColumnStart0('B')];
}

function getColumnB(inputRow, inputSheetCode) {
    switch (inputSheetCode){
        case 'IF':
            return inputRow[letterToColumnStart0('C')];
        case 'RD':
            return "Australian Tax Office";
        default:
            return null;
    }
}

function getColumnC(inputRow, inputSheetCode) {
    switch (inputSheetCode){
        case 'IF':
            return inputRow[letterToColumnStart0('E')];
        case 'RD':
            return inputRow[letterToColumnStart0('D')];
        default:
            return null;
    }
}

function getColumnD(inputRow, inputSheetCode) {
    switch (inputSheetCode){
        case 'IF':
            return inputRow[letterToColumnStart0('F')];
        case 'RD':
            return inputRow[letterToColumnStart0('E')];
        default:
            return null;
    }
}
function getColumnE(inputRow, inputSheetCode) {
    switch (inputSheetCode){
        case 'IF':
            return inputRow[letterToColumnStart0('H')];
        case 'RD':
            return inputRow[letterToColumnStart0('G')];
        default:
            return null;
    }
}
function getColumnF(inputRow, inputSheetCode) {
    switch (inputSheetCode){
        case 'IF':
            return inputRow[letterToColumnStart0('J')];
        case 'RD':
            return inputRow[letterToColumnStart0('I')];
        default:
            return null;
    }
}
function getColumnG(inputRow, inputSheetCode){
    var e = getColumnE(inputRow, inputSheetCode);
    var f = getColumnF(inputRow, inputSheetCode);
    if(!!e && !!f && f !== 0)
        return f/e;
    else
        return null;
}
function getColumnH(inputRow, inputSheetCode) {
    switch (inputSheetCode){
        case 'IF':
            return inputRow[letterToColumnStart0('L')];
        case 'RD':
            return inputRow[letterToColumnStart0('K')];
        default:
            return null;
    }
}
function getColumnI(inputRow, inputSheetCode) {
    switch (inputSheetCode){
        case 'IF':
            return inputRow[letterToColumnStart0('N')];
        case 'RD':
            return inputRow[letterToColumnStart0('M')];
        default:
            return null;
    }
}
function getColumnJ(inputRow, inputSheetCode) {
    if(getColumnL(inputRow, inputSheetCode) === 'NO')
        return "Not closed";
    else
        return "";
}
function getColumnK(inputRow, inputSheetCode) {
    switch (inputSheetCode){
        case 'IF':
            return inputRow[letterToColumnStart0('P')];
        case 'RD':
            return inputRow[letterToColumnStart0('O')];
        default:
            return null;
    }
}
function getColumnL(inputRow, inputSheetCode) {
    //TODO
    return "";
}
function getColumnM(inputRow, inputSheetCode) {
    switch (inputSheetCode){
        case 'IF':
            if(inputRow[letterToColumnStart0('X')] === 'No')
                return inputRow[letterToColumnStart0('S')] + inputRow[letterToColumnStart0('U')];
            else
                return inputRow[letterToColumnStart0('S')];
        case 'RD':
            return inputRow[letterToColumnStart0('R')] + inputRow[letterToColumnStart0('T')];
        default:
            return null;
    }
}
function getColumnN(inputRow, inputSheetCode) {
    //TODO
    return "";
}
function getColumnO(inputRow, inputSheetCode) {
    switch (inputSheetCode){
        case 'IF':
            return inputRow[letterToColumnStart0('W')];
        case 'RD':
            return inputRow[letterToColumnStart0('V')];
        default:
            return null;
    }
}
function getColumnP(inputRow, inputSheetCode) {
    //TODO
    return "";
}