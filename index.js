const express = require('express')
const app = express()
var fs = require('fs');
const cors = require('cors');
const moment = require("moment");
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('Worksheet Name');

const bodyparser = require('body-parser')
const axios = require('axios');

let reposNames=[];
const auth = {
            username: '010Devops',
            password: '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'
        };

async function getrepos() {
  try {
    const response = await axios.get(`https://api.github.com/user/repos`, { auth });
    const repos = response.data;
     return repos;
  } catch (error) {
    console.error(error);
  }
}

getrepos().then((repos) => {
  reposNames = repos.map(repo => repo.name?{name:repo.name,Update_at:repo.pushed_at}:'')
  console.log(reposNames,'repo');
  const headingColumnNames = [
      "Name",
      "Updated_at",
  ]
  //Write Column Title in Excel file
  let headingColumnIndex = 1;
  headingColumnNames.forEach(heading => {
      ws.cell(1, headingColumnIndex++)
          .string(heading)
  });
  //Write Data in Excel file
  let rowIndex = 2;
  reposNames.forEach( record => {
      let columnIndex = 1;
      Object.keys(record ).forEach(columnName =>{
          ws.cell(rowIndex,columnIndex++)
              .string(record [columnName])
              console.log(record,'rec',columnName);
      });
      rowIndex++;
  }); 

  date = moment().format("MM-DD-YYYY-T-HH-mm-ss");
  fileName = "D-"+ date;
  wb.write(`${fileName}.xlsx`);
});





