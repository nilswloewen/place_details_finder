import * as React from "react";
import { DefaultButton } from "office-ui-fabric-react";

export default class QueryColumnsTable extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      rows: []
    };
  }

  handleAddSelectedColumn = async () => {
    const selectedAddress = await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const selectedRange = context.workbook.getSelectedRange();
      const selectedColumns = selectedRange.getEntireColumn().load(["top", "columnIndex"]);
      await context.sync();
      const selectedColumn = sheet
        .getRangeByIndexes(selectedColumns.top, selectedColumns.columnIndex, 1, 1)
        .getEntireColumn()
        .load("address");
      await context.sync();
      return selectedColumn.address;
    });

    this.setState((prevState, props) => {
      if (!prevState.rows.includes(selectedAddress)) {
        return { rows: [...prevState.rows, selectedAddress] };
      }
    });
  };

  handleRemoveRow = () => {
    this.setState((prevState, props) => {
      return { rows: prevState.rows.slice(1) };
    });
  };

  render() {
    return (
      <div className="section">
        <div className="instructions">
          <span className="bullet">Step 2.</span>
          Select a column containing partial address data and add it to the list.
        </div>
        <table>
          <tbody id="query_columns_table">
            {this.state.rows.map((row, index) => (
              <tr key={index}>
                <td>{row}</td>
              </tr>
            ))}
          </tbody>
        </table>

        <DefaultButton className="query_btns" onClick={this.handleAddSelectedColumn}>Add</DefaultButton>
        {Boolean(this.state.rows.length) && <DefaultButton className="query_btns" onClick={this.handleRemoveRow}>Remove</DefaultButton>}
      </div>
    );
  }
}
