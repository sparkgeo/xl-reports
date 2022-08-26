XL Reports
----------

Generate Excel reports!

1. Create an Excel template file
2. Define a report configuration
3. Fetch your data
4. Generate report


## Report Configuration Schema

Report configuration is defined as array/list of objects/dicts.

```
"cell": <string> Worksheet cell coordinates to insert data, example: "B2"
"range": <string> Worksheet coordinate range to insert data. Range start coordinate is required and end coordinate is optional.
    examples: 
        * "B2:"
        * "B2:C5"
"data_key": <string> Key to use when fetching values from the data dictionary to insert into the worksheet. example: data["report_date"]
"sheet": <string> Worksheet name.
```

**Example configuration**

```
[
    {
        "cell": "B2",
        "data_key": "account",
        "sheet": "Area trends"
    },
    {
        "cell": "B4",
        "data_key": "report_date",
        "sheet": "Area trends"
    },
    {
        "range": "A8",
        "data_key": "report_data",
        "sheet": "Area trends"
    }
]
```

**Example data**

```
{
    "account": "Engineering"
    "report_date": str(date.today())
    "report_data": [
        [23.43, 11.96, 9.66],
        [6.99, 65.87, 45.33],
    ]
}
```