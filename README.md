
# Eclipse Recruitment Software Timesheet Converter

An application to convert the default 'export to xls.' timesheet file from
the popular recruitment software Eclipse to a more readable format for the
ability to import into the company payroll software.

Eclipse charge a fee to enable the report exporter options for each area of
their system, this wasn't viable for my use case so I needed to come up with
a way to take the default file and make it more readable for importing into
the company payroll software.

Their standard exporter is basically an excel screen reader of their output
to PDF. Meaning that the column formatting is split over multiple rows per
candidate with their data dotted around and not easily accessible.
## Features

As the plan was to provide this application to the payroll team directly and their lack of coding experience,
the application needed to run directly on their machines with minimal user input. Therefore the decision was
made to create a `tkinter` application that could be converted into a standalone `.EXE` file that they could
run as and when required (in this case weekly).


## Running the code locally

There are two ways to run the application locally for testing purposes.

#### Sample .EXE file

I have provided a sample `.EXE` file that can be run to test the system without having to compile the code yourself.

#### Running from the command line

Clone the project
```bash
  git clone https://github.com/csl619/eclipse_timesheet_converter
```
Go to the project directory
```bash
  cd eclipse_timesheet_converter
```
Install dependencies
```bash
  pip3 install -r requirements.txt
```
Run the application
```bash
  python3 file_converter.py
```
## Screenshots

![App Screenshot](https://imgur.com/28voeNG.png)
