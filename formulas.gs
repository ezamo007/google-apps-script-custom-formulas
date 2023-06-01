// Accepts a 1-column range and outputs a 2-column table showing the frequency of each item in the range. 
// Header can be disabled with the optional second argument.
function frequencies(values, includeHeader = true) {
    var header = ["Values", "Frequency"];
    var uniqueValues = Array.from(new Set(values.flat()));

    var frequencies = uniqueValues.map(function(value) {
        return [
            value, 
            values.flat().filter(function(item) {
                return item === value;
            }).length
        ];
    });

    output =  includeHeader ? [header, ...frequencies] : frequencies;
    return output
}

// Accepts a 2-column range and outputs a 3-column Venn diagram.
// Header can be disabled with the optional second argument.
function venn(values, includeHeader = true) {
    var column1 = values.map(row => row[0]);
    var column2 = values.map(row => row[1]);

    var onlyInColumn1 = column1.filter(item => !column2.includes(item));
    var bothColumns = column1.filter(item => column2.includes(item));
    var onlyInColumn2 = column2.filter(item => !column1.includes(item));

    var maxLength = Math.max(onlyInColumn1.length, bothColumns.length, onlyInColumn2.length);
    var output = includeHeader ? [["Left Only", "Both", "Right Only"]] : [];

    for (var i = 0; i < maxLength; i++) {
        var row = [
            i < onlyInColumn1.length ? onlyInColumn1[i] : "",
            i < bothColumns.length ? bothColumns[i] : "",
            i < onlyInColumn2.length ? onlyInColumn2[i] : ""
        ];
        output.push(row);
    }

    return output;
}

// Calculates URL of given cell address.
// By: Oluwaseun Olatoye
// From: https://www.oksheets.com/extract-hyperlink-url/
function url(input) {
    var myFormula = SpreadsheetApp.getActiveRange().getFormula();
    var myAddress = myFormula.replace('=GetURL(','').replace(')','');
    var myRange = SpreadsheetApp.getActiveSheet().getRange(myAddress);
    return myRange.getRichTextValue().getLinkUrl();
};


