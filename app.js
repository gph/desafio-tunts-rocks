const axios = require("axios");
const xl = require("excel4node");

const URL = "https://restcountries.com/v3.1/all";

const wb = new xl.Workbook();
const ws = wb.addWorksheet("Countries List");

const titleStyle = wb.createStyle({
    font: {
        bold: true,
        color: "#4F4F4F",
        size: 16,
    },
    alignment: {
        horizontal: "center"
    }
});

const headerStyle = wb.createStyle({
    font: {
        bold: true,
        color: "#808080",
        size: 12
    }
});

const numberStyle = wb.createStyle({
    numberFormat: '#.##0,0; (#.##0,0); -'
});

// Sheet Title
ws.cell(1, 1, 1, 4, true)
    .string("Countries List")
    .style(titleStyle);

ws.cell(2,1)
    .string("Name")
    .style(headerStyle);
ws.cell(2,2)
    .string("Capital")
    .style(headerStyle);
ws.cell(2,3)
    .string("Area")
    .style(headerStyle);
ws.cell(2,4)
    .string("Currencies")
    .style(headerStyle);

let rowCount = 3;

const countriesList = [];

axios.get(URL)
    .then(response => {
        const countries = response.data;
        countries.forEach(country => {
            countriesList.push(country.name.common = {
                name: country.name.common,
                capital: country.capital ? country.capital : "-",
                area: country.area ? country.area : "-",
                curriencies: country.currencies ? Object.keys(country.currencies).toString() : "-"
            });

            // // Name
            // ws.cell(rowCount,1).string(country.name.common);

            // // Capital
            // if (country.capital) {
            //     ws.cell(rowCount,2).string(country.capital[0]);
            // }else{
            //     ws.cell(rowCount,2).string("-");
            // }

            // // Area
            // if(country.area){
            //     ws.cell(rowCount,3)
            //         .number(country.area)
            //         .style(numberStyle);
            // }else{
            //     ws.cell(rowCount,3).string("-");
            // }

            // // Currency
            // if (country.currencies) {
            //     ws.cell(rowCount,4).string(Object.keys(country.currencies).toString());
            // } else {
            //     ws.cell(rowCount,4).string("-");
            // }
            // rowCount++;
        });
        countriesList.sort((a,b) => (a.name > b.name) ? 1 : ((b.name > a.name) ? -1 : 0))

        countriesList.forEach(country=>{
            console.log(country.name);
        });
        wb.write('countries-list.xlsx');
    })
    .catch(error => {
        console.log(error);
    });

