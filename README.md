# Excel-Outlook Email Search Tool

An innovative Excel VSTO Add-In designed to bridge Excel with Outlook, enabling users to perform searches within Outlook directly from Excel. This tool leverages cell values as search queries, significantly enhancing productivity and efficiency in managing large volumes of emails.

## Features

- **Excel Integration**: Seamlessly search Outlook emails using values from Excel cells.
- **Column-Based Searches**: Restricts searches to single-column selections to ensure precise query definitions.
- **Error Handling**: Incorporates comprehensive error handling to manage exceptions gracefully and guide users through correct tool usage.
- **Dynamic Column Creation**: Automatically adds columns in Excel to display search results, including email subjects and dates.

## Installation

This tool requires:
- Microsoft Excel and Outlook.
- .NET Framework 4.7.2.
- Visual Studio Tools for Office (VSTO) Runtime.

To install the Add-In:
1. Clone this repository or download the latest release.
2. Open the solution in Visual Studio.
3. Build the solution to generate the Add-In installer.
4. Run the installer to integrate the Add-In with Excel.
5. Or just download the [v2.0.zip](v2.0.zip) file containing the installer file to use it right away.

## Usage

1. Open Excel and select the cells containing your search queries. Ensure these cells are in a single column.
2. Click the "Search Email" button in the custom Ribbon tab.
3. If prompted, select an Outlook folder to search in.
4. Review the search results populated in new columns alongside your queries.

## Error Handling

The tool includes error handling to address common issues, such as:
- Invalid selections (e.g., multiple columns, empty cells).
- Unselected search folders.
- COM exceptions during Outlook item enumeration.

Error messages guide the user towards resolving these issues.

## Future Scope

- **Listing All Emails**: Option to list all emails found for each keyword in a dedicated Excel sheet or section.
- **Time Filters**: Introduce from and to date filters to narrow down the search period.
- **Improved UI Integration**: Enhance the Ribbon interface for more intuitive user interactions and feedback.
- **Advanced Query Options**: Support for complex queries combining multiple Excel columns.

## License

This project is licensed under the MIT License.

## Collaboration

Contributions, bug reports, and feature requests are welcome.
