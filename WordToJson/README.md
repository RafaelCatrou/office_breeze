# Word To Json

### Table of Contents
**[Goal](#goal)**
**[Demo](#demo)**
**[Note](#note)**

## Goal

Provide a mean to extract data from Content Controls in Word, and export data to JSON.

The release is in the folder: *WordToJson_NET_4_5_2\release\version\WordToJson_NET_4_5_2.exe*

## Demo

### Execute provided .exe (in release)

```
WordToJson_NET_4_5_2.exe Demo.docx
```

### Results

In the same folder, a JSON file called *Demo.docx.json* is generated:

```
{
  "path": "d:\\GitHub\\office_breeze\\WordToJson\\Demo.docx.json",
  "Example 1": "Value 1",
  "Example 2": [
    "Value 2",
    "Value 3",
    "Value 4",
    ]
}
```

## Note

Microsoft website: [More Info on Interop](https://docs.microsoft.com/fr-fr/dotnet/csharp/programming-guide/interop/how-to-access-office-onterop-objects)
