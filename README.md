# TemplatedExcel
Defining Excel styles with HTML and CSS. It's a templated language, as same as HTML.
# News and noteworthy
- V1.1.0 release 2019-10-04
	- Adjust the 'autoColumnSize' after the content generated.
	- Create the style of cell lazyly.
- V1.0.1 release 2019-04-28
	- Providing the basic capbility for excel generating, includes CSS defined cell style, dynamic template with Java objects which can be a getter method or Hash.
	- Interface abstracted for various  Excel adapter

# Template language tags
|Tag Name |Attribute   |Description   |
| ------------ | ------------ | ------------ |
|workbook   |--|The root element   |
|sheet   |name   |The name of sheet, it will be displayed in Excel   |
|row   |height   |The height of row   |
|   |style   |The style of the row will apply to all cells included it    |
|cell   |fit-content   |Adjusts the column width to fit the contents.    |
|   |quote-prefixed   |Let numbers appear as non-numeric   |
|   |colspan   |Allows a single excel cell to span the width of more than one cell or column.   |
|   |rowspan   |Allows a single excel cell to span the height of more than one cell or row.   |
|   |style   |The style of cell which describes Excel cell style   ||

# CSS properties supported
`Not supported by all CSS standards, all color is hex color definition only.`
- ***background-color***: Hex color definition supported only.  eg: #FFFFFF.
- ***text-color***: Hex color definition supported only.  eg: #FFFFFF.
- ***horizontal-alignment***: left, right, center supported only.
- ***vertical-aligment***: top, bottom, center supported only.
- ***border***: eg: 1px thin #000000, The border-style supported is thin, medium, dashed, dotted, thick, double and hair only.
- ***border-top-style***: as same as border.
- ***border-bottom-style***: as same as border.
- ***border-left-style***: as same as border.
- ***border-right-style***: as same as border.
- ***border-top-color***: as same as border.
- ***border-bottom-color***: as same as border.
- ***border-left-color***: as same as border.
- ***border-right-color***: as same as border.
- ***font-weight***: bold or others.
- ***font-family***: eg: Microsoft YaHei, the font name in excel.
- ***font-size***: eg: 12px, the font size in excel.
- ***font-style***: italic or others.
- ***text-decoration***: underline or others.
# Reference
[ph-css](https://github.com/phax/ph-css "ph-css") For parsing CSS text,[thymeleaf](https://github.com/thymeleaf/thymeleaf "thymeleaf") For dynamic template file.

# Quick Start
## Maven dependency
```xml
<dependency>
  <groupId>com.github.braisdom</groupId>
  <artifactId>templated-excel</artifactId>
  <version>1.1.0</version>
</dependency>
```
To build TemplatedExcel from source, Maven 3.0.4 is required. Any Maven version below does NOT work!

## Example
Java Code:
```java
public class Sample {

    public static void main(String[] args) throws Exception {
        InputStreamReader inputStreamReader = new InputStreamReader(Sample
                .class.getResourceAsStream("/template.xml"));
        File excelFile = new File("./sample.xls");
        WorkbookTemplate workbookTemplate = new WorkbookTemplate(inputStreamReader);
        workbookTemplate.process(new SampleTemplateDataSource(), new PoiWorkBookWriter(), excelFile);
    }
}
```
Template Code:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:th="https://www.braisdom.org/templated-excel">
    <sheet name="StyleSheet">
        <row height="40">
            <cell fit-content="true" quote-prefixed="true">00000000012344</cell>
            <cell fit-content="true" colspan="2" style="text-align: center;vertical-align: center;">
                Merged Column, Text Align Center, Vertical Align Center
            </cell>
        </row>
        <row>
            <cell fit-content="true" style="color: #FF0000;">Color Text</cell>
        </row>
        <row>
            <cell fit-content="true"
                style="background-color: #FF0000;color: #FFFFFF;">Background Color</cell>
        </row>
        <row>
            <cell fit-content="true" style="border:1px thin #FF0000;">Cell Border</cell>
        </row>
        <row>
            <cell style="font-style: italic;font-weight: bold;
                        font-family: Microsoft YaHei;text-decoration: underline;">
                Font Style
            </cell>
        </row>
        <row height="50">
            <cell fit-content="true" style="border:1px thin #FF0000;">Row Height</cell>
        </row>
    </sheet>
    <sheet name="DataTableSheet">
        <data-table>
            <header>
                <row height="30">
                    <cell colspan="2" style="font-size: 18;font-weight: bold;
                            text-align: center;vertical-align: center;">
                        Employee Table
                    </cell>
                </row>
                <row>
                    <cell fit-content="true" style="font-weight: bold;
                            border: 1px thin #000000;">Name</cell>
                    <cell fit-content="true" style="font-weight: bold;
                            border: 1px thin #000000;">Gender</cell>
                    <cell fit-content="true" style="font-weight: bold; s
                            border: 1px thin #000000;">Occupation</cell>
                </row>
            </header>
            <body>
                <row th:each="user : ${users}">
                    <cell fit-content="true" style="border: 1px thin #000000;"
                          th:text="${user.name}"/>
                    <cell fit-content="true" style="border: 1px thin #000000;"
                          th:text="${user.gender}"/>
                    <cell fit-content="true" style="border: 1px thin #000000;"
                          th:text="${user.occupation}"/>
                </row>
            </body>
        </data-table>
    </sheet>
</workbook>
```
![](https://raw.githubusercontent.com/braisdom/TemplatedExcel/master/images/style.png)
![](https://github.com/braisdom/TemplatedExcel/blob/master/images/data-table.png?raw=true)
