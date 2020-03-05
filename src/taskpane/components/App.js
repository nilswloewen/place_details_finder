import * as React from "react";
import Progress from "./Progress";
import QueryColumnsTable from "./QueryColumnsTable";
import GooglePlacesApi from "./GooglePlacesApi";
import InitOutputRangeBtn from "./InitOutputRangeBtn";
import BuildJsonBtn from "./BuildJsonBtn";

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
  }

  onSelectionChange = async args => {
    console.log("onSelectionChange() fired.");
    let get = async () => {
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(args.worksheetId);
        const selectedRange = sheet.getRange(args.address).load(["address", "rowCount", "rowIndex"]);
        await context.sync();

        const query = await this.buildQueryFromRowIndex(context, args.worksheetId, selectedRange.rowIndex);

        // Display selected range and amount of rows.
        document.getElementById("selected_address").innerText = args.address + 1;
        document.getElementById("rows_selected").innerText = selectedRange.rowCount;
        document.getElementById("selected_row_index").innerText = selectedRange.rowIndex;
        document.getElementById("query_input").value = query.join(" ");
      });
    };
    get.bind(args);
    get();
  };

  getQueryColumnAddresses = () => {
    const table = document.getElementById("query_columns_table");
    let addresses = [];
    for (let i = 0; i < table.rows.length; i++) {
      addresses.push(table.rows[i].cells[0].innerText);
    }
    return addresses;
  };

  buildQueryFromRowIndex = async (context, worksheetId, rowIndex) => {
    const addresses = this.getQueryColumnAddresses();
    let queryValues = [];

    for (const address of addresses) {
      const val = await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(worksheetId);
        const column = sheet.getRange(address).load("columnIndex");
        await context.sync();
        const cell = sheet.getRangeByIndexes(rowIndex, column.columnIndex, 1, 1).load("values");
        await context.sync();
        return cell.values[0][0];
      });
      queryValues.push(val);
    }

    return queryValues;
  };

  attachSelectionEventToTable = async () => {
    Excel.run(async context => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onSelectionChanged.add(this.onSelectionChange);

      await context.sync();
      console.log("onSelectionChanged event successfully registered SourceTable.");
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return <Progress title={title} message="Details Finder is loading..." />;
    }

    this.attachSelectionEventToTable();

    return (
      <div>
        <table>
          <tbody>
            <tr>
              <td>Selected Address</td>
              <td id="selected_address" />
            </tr>

            <tr>
              <td>Selected Row Index</td>
              <td id="selected_row_index" />
            </tr>

            <tr>
              <td>Number of rows</td>
              <td id="rows_selected" />
            </tr>

            <tr>
              <td colSpan="2">
                <QueryColumnsTable />
              </td>
            </tr>

            <tr>
              <td>Query</td>
            </tr>
            <tr>
              <td colSpan="2">
                <input type="textarea" id="query_input" placeholder={"Click on a row..."} style={{ width: "280px" }} />
              </td>
            </tr>
          </tbody>
        </table>

        <InitOutputRangeBtn />
        <GooglePlacesApi />
        <BuildJsonBtn />
      </div>
    );
  }
}
