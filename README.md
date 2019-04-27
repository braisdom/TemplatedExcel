# TemplatedExcel
Defining Excel styles with HTML and CSS. It's a templated language, as same as HTML.
# News and noteworthy
- V1.0.0 release 2019-04-28
	- Providing the basic capbility for excel generating, includes CSS defined cell style, dynamic template with Java objects which can be a getter method or Hash.
	- Interface abstracted for various  Excel adapter

# Quick Start
## Maven dependets
```xml
<dependency>
  <groupId>com.github.braisdom</groupId>
  <artifactId>templated-excel</artifactId>
  <version>1.0.0</version>
</dependency>
```
To build TemplatedExcel from source, Maven 3.0.4 is required. Any Maven version below does NOT work!

## Example
Java Code:
```java
public class Sample {

    public static void main(String[] args) throws Exception {
        InputStreamReader inputStreamReader = new InputStreamReader(Sample.class.getResourceAsStream("/template.xml"));
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
            <cell fit-content="true" style="background-color: #FF0000;color: #FFFFFF;">Background Color</cell>
        </row>
        <row>
            <cell fit-content="true" style="border:1px thin #FF0000;">Cell Border</cell>
        </row>
        <row>
            <cell style="font-style: italic;font-weight: bold;font-family: Microsoft YaHei;text-decoration: underline;">
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
                    <cell colspan="2" style="font-size: 18;font-weight: bold;text-align: center;vertical-align: center;">
                        Employee Table
                    </cell>
                </row>
                <row>
                    <cell fit-content="true" style="font-weight: bold;border: 1px thin #000000;">Name</cell>
                    <cell fit-content="true" style="font-weight: bold;border: 1px thin #000000;">Gender</cell>
                    <cell fit-content="true" style="font-weight: bold;border: 1px thin #000000;">Occupation</cell>
                </row>
            </header>
            <body>
                <row th:each="user : ${users}">
                    <cell fit-content="true" style="border: 1px thin #000000;" th:text="${user.name}"/>
                    <cell fit-content="true" style="border: 1px thin #000000;" th:text="${user.gender}"/>
                    <cell fit-content="true" style="border: 1px thin #000000;" th:text="${user.occupation}"/>
                </row>
            </body>
        </data-table>
    </sheet>
</workbook>
```
Styled Excel:
![](https://raw.githubusercontent.com/braisdom/TemplatedExcel/master/images/style.png)

Dynamic Excel:
![](https://github.com/braisdom/TemplatedExcel/blob/master/images/data-table.png?raw=true)