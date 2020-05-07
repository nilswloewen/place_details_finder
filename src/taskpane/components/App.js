import * as React from "react";
import Progress from "./Progress";
import QueryColumnsTable from "./QueryColumnsTable";
import GooglePlacesApi from "./GooglePlacesApi";
import InitOutputRangeBtn from "./InitOutputRangeBtn";
import BuildJsonBtn from "./BuildJsonBtn";
import ApiKeyForm from "./ApiKeyForm";

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      apiKey: null
    };
  }

  componentDidMount = () => {
    OfficeRuntime.storage.getItem("apiKey").then(key => {
      this.setState({
        apiKey: key
      });
    });
  };

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
        document.getElementById("query_input").innerText = query.join(" ");
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

    if (this.state.apiKey === null) {
      return <ApiKeyForm />;
    }

    this.attachSelectionEventToTable();

    return (
      <div>
        <div id="selected_address" className="hidden" />
        <div id="selected_row_index" className="hidden" />
        <div id="rows_selected" className="hidden" />
        <InitOutputRangeBtn />
        <QueryColumnsTable />
        <div className="section">
          <div className="instructions">
            <span className="bullet">Step 3.</span>
            Review and modify the query built from your selection.
          </div>
          <div contentEditable={true} id="query_input" placeholder={"Click on a row..."} style={{ width: "280px" }} />
        </div>
        <GooglePlacesApi apiKey={this.state.apiKey} />
        <ApiKeyForm apiKey={this.state.apiKey} />
        <BuildJsonBtn />
      </div>
    );
  }
}
