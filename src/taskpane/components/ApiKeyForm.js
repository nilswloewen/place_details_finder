import * as React from "react";

export default class ApiKeyForm extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      apiKey: this.props.apiKey
    };

    this.handleChange = this.handleChange.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
  }

  handleChange(event) {
    this.setState({ apiKey: event.target.value });
  }

  storeKey = async value => {
    const key = "apiKey";
    console.log("storeKey:" + value);
    const report = await OfficeRuntime.storage.setItem(key, value).then(
      function(result) {
        return "Success: Item with key '" + key + "' saved to storage.";
      },
      function(error) {
        return "Error: Unable to save item with key '" + key + "' to storage. " + error;
      }
    );
    console.log(report);
  };

  handleSubmit = async event => {
    event.preventDefault();
    await this.storeKey(this.state.apiKey);
    window.location.reload(false);
    // this.checkForValidApiKey();
  };

  checkForValidApiKey() {
    const err = document.getElementById("google_error");
    if (err.innerText === "GoogleError") {
      this.setState({ apiKey: "Invalid" });
    }
  }

  render() {
    return (
      <div className="section">
        <div className="instructions">
          <span className="bullet">Link with Google Places API</span>
          Enter your <a href="https://cloud.google.com/maps-platform/">API key</a>.
        </div>
        <label>
          ApiKey:
          <input defaultValue={this.state.apiKey} onChange={this.handleChange} placeholder="Paste Key Here" />
        </label>
        <input type="submit" value="Submit" onClick={this.handleSubmit} />
      </div>
    );
  }
}
