import * as React from "react";

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
      <div>
        <table>
          <thead>
            <tr>
              <th>Build query from these columns</th>
            </tr>
          </thead>
          <tbody id="query_columns_table">
            {this.state.rows.map((row, index) => (
              <tr key={index}>
                <td>{row}</td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr>
              <td onClick={this.handleAddSelectedColumn}>(+) Add selected column</td>
              {Boolean(this.state.rows.length) && <td onClick={this.handleRemoveRow}>(-)</td>}
            </tr>
          </tfoot>
        </table>
      </div>
    );
  }
}
