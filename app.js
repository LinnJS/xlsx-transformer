const xlsx = require('xlsx');

const options = { cellDates: true };
const workbook = xlsx.readFile('VISN-17_Facility_HS_before.xlsx', options);
const worksheet = workbook.Sheets['Sheet1'];

// turn worksheet data into JSON so we can modify the data
const JSONData = xlsx.utils.sheet_to_json(worksheet);

// remove rows that include N/A
const removeNA = JSONData.filter((record) => !record['ID'].includes('n/a'));
// remove rows that include "COVID-19 Vaccines"
const JSONDataWithoutCOVID = removeNA.filter((record) => !record['ID'].includes('COVID-19'));

// TODO: change your environmental delimiter on your system from “,” to “;”

const removeAndAddRows = JSONDataWithoutCOVID.map((record) => {
  /* Data shape of record
    {
     "ID": "vha_740GC - Pharmacy",
     "Facility ID": "vha_740GC",
     "Facility": "Corpus Christi VA Clinic",
     "VACM System Health Service": "Pharmacy at VA Texas Valley health care",
     "Health Services": "Pharmacy",
     "Health System": "VA Texas Valley health care",
     "Owner": "VA Texas Valley health care"
  }*/

  // delete column B = Facility ID
  delete record['Facility ID'];
  // delete column E = Health Services
  delete record['Health Services'];
  // delete column F = Health system
  delete record['Health System'];
  // keep columns ID, facility, VAMC System Health Service, Facility description, owner

  // if there is not a Facility description add a blank one
  const hasFacilityDescription = !!record['Facility description of services'];
  hasFacilityDescription ? record['Facility description of services'] : (record['Facility description of services'] = '');

  return record;
});

// TODO: remove duplicates from column A (string match?) what is column A? ID?
const removeDuplicateRecords = removeAndAddRows.reduce((acc, record) => {
  const hasDuplicate = (row) => {
    return row.ID === record.ID;
  };

  if (!acc.some(hasDuplicate)) {
    acc.push(record);
  }

  return acc;
}, []);

const removeAposFromRecords = removeAndAddRows.map((record) => {
  const recordKeys = Object.keys(record);

  const newData = recordKeys.reduce((acc, key) => {
    const value = record[key];
    const cleanedValue = value.replace(/'/g, '');

    return {
      ...acc,
      [key]: cleanedValue,
    };
  }, {});

  return newData;
});

const toDoubleQuotedJSON = (json) => {
  const JSONString = JSON.stringify(json);
  const JSONWithDoubleQuotes = JSONString.replace(/'/g, '"');

  return JSON.parse(JSONWithDoubleQuotes);
};

// create workbook
const newWorkbook = xlsx.utils.book_new();
// Add transformed data to new worksheet
const newWorksheet = xlsx.utils.json_to_sheet(toDoubleQuotedJSON(removeAposFromRecords));
// append worksheet to workbook
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Transformed facilities list');
// print that workbook to the project directory
xlsx.writeFile(newWorkbook, 'transformedFile.xlsx');

// create csv
const newCSV = xlsx.utils.sheet_to_csv(newWorksheet, {
  FS: ';',
  RS: '\n',
  dateNF: 'dd"."mm"."yyyy',
  strip: false,
  blankrows: true,
});

console.log(newCSV);
