import JSZip from "jszip";
import { workbookBufferToRows } from "../src/xlsx.js";

describe("workbookBufferToRows", () => {
  it("reads lightweight generated workbooks with prefixed spreadsheet XML", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="utf-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
        <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml" />
      </Types>`
    );
    zip.file(
      "_rels/.rels",
      `<?xml version="1.0" encoding="utf-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" Id="rId1" />
      </Relationships>`
    );
    zip.file(
      "xl/workbook.xml",
      `<?xml version="1.0" encoding="utf-8"?>
      <x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:sheets>
          <x:sheet name="BS_Data_May26" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
        </x:sheets>
      </x:workbook>`
    );
    zip.file(
      "xl/_rels/workbook.xml.rels",
      `<?xml version="1.0" encoding="utf-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" Id="rId1" />
      </Relationships>`
    );
    zip.file(
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0" encoding="utf-8"?>
      <x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:sheetData>
          <x:row r="1">
            <x:c r="A1" t="str"><x:v>Account label</x:v></x:c>
            <x:c r="B1" t="str"><x:v>Date</x:v></x:c>
            <x:c r="C1" t="str"><x:v>Amount</x:v></x:c>
          </x:row>
          <x:row r="2">
            <x:c r="A2" t="str"><x:v>Cash</x:v></x:c>
            <x:c r="B2" t="n"><x:v>46173</x:v></x:c>
            <x:c r="C2" t="n"><x:v>123.45</x:v></x:c>
          </x:row>
        </x:sheetData>
        <x:tableParts count="1"><x:tablePart r:id="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" /></x:tableParts>
      </x:worksheet>`
    );
    zip.file(
      "xl/worksheets/_rels/sheet1.xml.rels",
      `<?xml version="1.0" encoding="utf-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="/xl/tables/table1.xml" Id="rId2" />
      </Relationships>`
    );
    zip.file(
      "xl/tables/table1.xml",
      `<?xml version="1.0" encoding="utf-8"?>
      <x:table id="1" name="BS_Data_May26" displayName="BS_Data_May26" ref="A1:C2" headerRowCount="1" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:tableColumns count="3">
          <x:tableColumn id="1" name="Account label" />
          <x:tableColumn id="2" name="Date" />
          <x:tableColumn id="3" name="Amount" />
        </x:tableColumns>
      </x:table>`
    );

    const rows = await workbookBufferToRows(Buffer.from(await zip.generateAsync({ type: "nodebuffer" })));

    expect(rows).toEqual([{ "Account label": "Cash", Date: 46173, Amount: 123.45 }]);
  });
});
