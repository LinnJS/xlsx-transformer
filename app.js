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

  // TODO: remove duplicates from column A (string match?) what is column A? ID?
  return record;
});

const removeAposFromRecords = removeAndAddRows.map((record) => {
  const recordKeys = Object.keys(record);

  recordKeys.map((key) => {
    const value = record[key];

    if (value.includes("'")) {
      value.replace(/\W/g, ' ');
      console.log('value: ', value);
      return value;
    }
  });

  return record;
});

console.log('removeAposFromRecords: ', removeAposFromRecords);

const toDoubleQuotedJSON = (json) => {
  const JSONString = JSON.stringify(json);
  const JSONWithoutApos = JSONString.replace(/'/, '');
  const JSONWithDoubleQuotes = JSONWithoutApos.replace(/'/g, '"');

  return JSON.parse(JSONWithDoubleQuotes);
};
// toDoubleQuotedJSON(removeAndAddRows);

// console.log('removeAndAddRows: ', removeAndAddRows);
// console.log('toDoubleQuotedJSON: ', toDoubleQuotedJSON(removeAndAddRows));

const newWorkbook = xlsx.utils.book_new();
const newWorksheet = xlsx.utils.json_to_sheet(removeAndAddRows);
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Transformed facilities list');

xlsx.writeFile(newWorkbook, 'transformedFile.xlsx');
