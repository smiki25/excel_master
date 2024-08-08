# Excel Master Flask

This is a Flask application that allows you to update an Excel file with new data and retrieve the updated information.

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/your-username/excel-master-flask.git
    ```

2. Install the required dependencies:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. Start the Flask server:

    ```bash
    python excel_master_flask.py
    ```

2. Send a POST request to `http://<here_should_be_the_url>/excel_master/update_excel` following this JSON payload template:

    ```json
    {
      "interest_y": "3.00%",
      "number_of_periods": 500,
      "principal": 260000
    }
    ```

    This will update the Excel file with the new data.

3. The server will respond with the updated information in JSON format:

    ```json
    {
      "payment": 1234.56,
      "interest_y": "3.00%",
      "number_of_periods": 500,
      "principal": 260000
    }
    ```
