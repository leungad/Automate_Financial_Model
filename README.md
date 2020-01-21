# automate_fin_model
Automate Financial Modeling from Capital IQ

# Automate_Financial_Model

Utilizing OpenPyXL to automate financial modeling tasks within equity research. The files that are converted/formatted are standardized Capital IQ Financial Statements/ Key Statistics. The script will firstly generate a generic revenue model from the revenue segments, a discounted cashflow model from the income statement, and a sheet for the discounted cashflow assumptions. It leaves empty the percentages for the user to input their assumtions of growth.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

An example of an Excel file will be attached. 

### Prerequisites

What things you need to install the software and how to install them

```
Give examples
```

### Installing

A step by step series of examples that tell you how to get a development env running

Say what the step will be

```
Give the example
```
![Image of Excel File](https://picturesadrianblog.s3-us-west-2.amazonaws.com/BTS_Pic.png)
Visual of standardized XLSX file generated from the Capital IQ database.

![Image of Revenue Model Output](https://picturesadrianblog.s3-us-west-2.amazonaws.com/revenue_model.png)
Once formatted the generated output is shown. It can be customized by color for historical and forecasted colors (Blue/ Light Blue)
I will be implementing an option to input what colors you want to specify.

```
until finished
```

I plan to generate a web-application that will allow users to input the files, and generate a download of the output.

## Running the tests

Explain how to run the automated tests for this system

### Break down into end to end tests

Explain what these tests test and why

```
Give an example
```

### And coding style tests

Explain what these tests test and why

```
Give an example
```

## Deployment

Add additional notes about how to deploy this on a live system

## Built With

* [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/index.html) - The excel manipulation library used

## Contributing

Please read [CONTRIBUTING.md](https://gist.github.com/PurpleBooth/b24679402957c63ec426) for details on our code of conduct, and the process for submitting pull requests to us.

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 

## Authors

* **Adrian Leung** - *Initial work* - (https://github.com/leungad)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Hat tip to anyone whose code was used
* Inspiration
* etc
