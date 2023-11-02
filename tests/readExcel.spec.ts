import { test, expect } from "@playwright/test";
import XLSX from "xlsx";



test("Read a Simple Excel File", async ({ page }) => {
    const org_path = "original.xlsx";
    var workbook = XLSX.readFile(org_path)

    let worksheet= workbook.Sheets[workbook.SheetNames[0]];

    for(let index=2; index<7;index++){
        let firstColVal = worksheet[`A${index}`].v;
        const secondColVal = worksheet[`B${index}`].v;
        const thirdColVal = worksheet[`C${index}`].v;
        const fourthColVal = worksheet[`D${index}`].v;
        const fifthColVal = worksheet[`E${index}`].v;
        const sixthColVal = worksheet[`F${index}`].v;

        firstColVal= new Date((firstColVal - (25567 + 1))*86400*1000);

        console.log(firstColVal,  secondColVal,  thirdColVal,  fourthColVal,  fifthColVal,  sixthColVal)
    }




})