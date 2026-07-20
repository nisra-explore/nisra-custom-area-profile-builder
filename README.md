# #81 DAERA Rural Dashboard

## Table of Contents

- [#81 DAERA Rural Dashboard](#daera-rural-dashboard)
  - [Table of Contents](#table-of-contents)
  - [:newspaper: Aim](#newspaper-aim)
  - [:house: Structure](#house-structure)
    - [File structure](#file-structure)
    - [Data Input](#data-input)
    - [Code structure](#code-structure)
    - [Software Checklist](#software-checklist)
      - [Git set up](#git-set-up)
  - [:arrows\_clockwise: Processes](#arrows_clockwise-processes)
    - [Process Diagram](#process-diagram)
    - [:information\_source: Indicator sources](#information_source-indicator-sources)
      - [Updating an indicator](#updating-an-indicator)
      - [Adding a new indicator/domain](#adding-a-new-indicatordomain)
    - [Link with Data Portal](#link-with-data-portal)
    - [Update the dashboard with any commentary on trends :chart\_with\_upwards\_trend:](#update-the-dashboard-with-any-commentary-on-trends-chart_with_upwards_trend)
    - [Process for updating code](#process-for-updating-code)
    - [Testing phase :mortar\_board:](#testing-phase-mortar_board)
    - [explore.nisra.gov.uk hosting :computer:](#explorenisragovuk-hosting-computer)
    - ['Live' check :sun\_with\_face:](#live-check-sun_with_face)
  - [:warning: Troubleshooting](#warning-troubleshooting)
  - [Frequently Asked Questions](#frequently-asked-questions)
    - [How do we add a new page, e.g. notes?](#how-do-we-add-a-new-page-eg-notes)
    - [How do we hide pages?](#how-do-we-hide-pages)
    - [How do we change branding, logos, etc.?](#how-do-we-change-branding-logos-etc)
    - [How do we change colours of chart, maps, boxes?](#how-do-we-change-colours-of-chart-maps-boxes)
    - [How do we change chart styles?](#how-do-we-change-chart-styles)
    - [What parts of the script do we need to update if we move to the live data portal?](#what-parts-of-the-script-do-we-need-to-update-if-we-move-to-the-live-data-portal)
    - [The process of updating GitHub when we make changes?](#the-process-of-updating-github-when-we-make-changes)
    - [What's the process for publishing the dashboard?](#whats-the-process-for-publishing-the-dashboard)
  - [:question: Links](#question-links)

## :newspaper: Aim
Documentation to outline the structure and processes needed to create or modify the DAERA Rural dashboard.

## :house: Structure

### File structure 

| File | Purpose  |
| --- | --- |
| `index.html` | The dashboard main page |
| `style.css` | Pre-defined styling for the dashboards - colours, fonts, sizing, spacing etc. |
| `config.R` | Load the packages and functions required for data import and preparation |
| `read_all_dataportal_tables_in.R` | Importing and preparing the data from the Flexible Table Builder |
| `read_all_flexible_table_builder_data.R` | Importing and preparing the data from the NISRA Data Portal |
| `create_data.R` | Runs the other three R scripts and creates the final JSON file which populates the dashboard |

### Software Checklist

- Visual Studio Code (with "Live Server" Extension)
- R Studio
- Git for Windows
 
#### Git set up

If you are using Git for the first time follow these configuration steps before cloning the Git Repository:
1. Register an account on github.com with your work email address.
2. Open Visual Studio Code
3. Open a new Terminal, either by clicking `Terminal` in the top menu and choosing `New Terminal` or pressing <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>'</kbd>
4. In the Terminal pane enter each line of code, pressing <kbd>Enter</kbd> __after each line__:
    -  `git config --global http.sslVerify false`
    - `git config --global http.proxy http://cloud-lb.nigov.net:8080`
    - `git config --global https.proxy https://cloud-lb.nigov.net:8080`
    - `git config --global user.name "YourUserName"`
    - `git config --global user.email first.last@nisra.gov.uk`

You will need to enter the username and email address you registered your github.com account with. 

## :arrows_clockwise: Processes

### Process Diagram

The diagram below shows how the functionality behind this dashboard renders all the code. Some parts of the process are _independent_ (they occur automatically as the page loads) and some are _dependent_ (they occur in response to some interaction from the user).

<div style="width: 100%;">
  <img src="img/data-flow-chart.drawio (1).svg" style="width: 100%;" alt="Click to see the source">
</div>

### Code structure

#### Data Input

#### index.html

##### File structure

| Section | Purpose  |
| --- | --- |
| Head | Page title, DAERA logo, import css and js dependencies, controls zoom and layout for different screen sizes |
| html content | Defines the tables, charts, map and buttons; adds styling |
| Javascript content | Functions for dynamically creating theinteractive map, tables, chart, area profile and data downloads |

### Process for Adding or Removing a Variable

1. Open R Studio
2. Run a "Git pull" to ensure code is up to date with repository
3. Update the 'table_list' to add or remove the variable name
4. Update the 'further_breakdown_df' to add the category name for the url link
5. Save the changes and run the 'create_data.R' script
6. In the 'index.html' file, go to the 'category-selector' div which contains all of the category cehckboxes 
7. Remove the checkbox to remove a variable, or copy and paste a new checkbox with the name of the variable as the 'Value' and add a label to the checkbox

    ```
      <label><input type="checkbox" value="Household Tenure" checked> Tenure</label>
    ```

### Changing colours of chart and maps?
 * For chart colours, see the `chart_config` definition inside the `createLineChart()` function in the [`data_functions.js`](scripts/data_functions.js) script. See [Chart.js documentation](https://www.chartjs.org/docs/latest/) on ways to make changes.
 * For map colours, see the `drawMap()` function in the [`data_functions.js`](scripts/data_functions.js) script. See [leaflet.js documentation](https://leafletjs.com/reference.html) on how to customise maps.
 * For boxes, see the [`style.css`](style.css) stylesheet. Find the corresponding id or class of the page element you wish to change and change the `background-color` property.

### The process of updating GitHub when we make changes?
See [Process for updating code](#process-for-updating-code).

## :question: Links
- [Chart.js documentation](https://www.chartjs.org/docs/latest/)
- [Chart.js YouTube tutorials](https://www.youtube.com/c/ChartJS-tutorials)
- [CSS / styling guide on W3schools](https://www.w3schools.com/Css/)
- [HTML guide on W3schools](https://www.w3schools.com/html/default.asp)
- [Javascript guide on W3schools](https://www.w3schools.com/js/default.asp)
- [Free introductory html/css/js courses](https://www.codecademy.com/)
- [Test HTML/css/javascript on web sandbox - DO NOT UPLOAD SENSITIVE DATA](https://jsfiddle.net/)
- [Visual Studio Code documentation](https://code.visualstudio.com/Docs)
- [Github.com documentation](https://docs.github.com/en)
- [R cheatsheet PDFs](https://github.com/rstudio/cheatsheets)
