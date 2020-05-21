import * as React from "react";
import Progress from "./Progress";
import QueryColumnsTable from "./QueryColumnsTable";
import GooglePlacesApi from "./GooglePlacesApi";
import InitOutputRangeBtn from "./InitOutputRangeBtn";
import BuildJsonBtn from "./BuildJsonBtn";

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      apiKey: ""
    };
    this.handleChange = this.handleChange.bind(this);
  }

  componentDidMount = async () => {
    if (!this.state.validKey) {
      setTimeout(function() {}, 5000);
      const error = document.getElementById("google_error");
      if (error) {
        if (error.innerText === "NoErrorReported") {
          this.setState({
            validKey: true
          });
        }
      }
    }
  };

  handleChange(event) {
    let input = event.target.value;
    if (input) {
      this.setState({ apiKey: input.trim() });
    }
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

  checkForValidApiKey() {
    const err = document.getElementById("google_error");
    if (err.innerText === "GoogleError") {
      this.setState({ apiKey: "Invalid", validKey: false });
    }
  }

  render() {
    const { title, isOfficeInitialized } = this.props;
    if (!isOfficeInitialized) {
      return <Progress title={title} message="Details Finder is loading..." />;
    }

    if (!this.state.apiKey || !this.state.validKey) {
      return (
        <div className="section">
          <div className="instructions">
            <span className="bullet">Link with Google Places API</span>
            Enter your <a href="https://cloud.google.com/maps-platform/">API key</a>.
          </div>
          <label>
            <input defaultValue={this.state.apiKey} onChange={this.handleChange} placeholder="Paste Key Here" />
          </label>
        </div>
      );
    }

    this.attachSelectionEventToTable();

    return (
      <div>
        <div id="selected_address" className="hidden" />
        <div id="selected_row_index" className="hidden" />
        <div id="rows_selected" className="hidden" />
        <label>
          API Key:
          <div id="apiKey">{this.state.apiKey}</div>
        </label>
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
        <BuildJsonBtn />
      </div>
    );
  }
}
