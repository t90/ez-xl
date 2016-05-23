# EZ-XL easy Excel
If you need a simple and lightweight solution to help you render your data into an excel multipage spreasheet with some formatting and formulas. Look no further.

This one class library will allow you to create .ODS files which are natively supported by MS Excel 2013 (c) (tm) and higher. 

All you need to do is prepare your excel template. Mark the cells with %PropertyName markers, save it as ODS in excel and 
feed that template to the library and let it merge excel with your data.

## Save your excel file as .ODS

- In your Excel ribbon click File
- Save As
- Computer
- Pick a folder (like "My Documents" or "c:\work")
- Open "Save as type:" dropdown
- Select "OpenDocument Spreadsheet"
- Click "Save"
- Click "Yes" for the prompt

Let's say you saved file in "c:\work" as "template.ods"

And the template you have looks like (it has two pages called 'Sheet1' and 'Sheet2')

```
Sheet1
 |    A      |    B     |   
1|UserName   |    Fee   |
2|%UserName  |%Fee      |

Sheet2
  |    A      |    B              |
1 |UserName   |    Service Date   |
2 |%UserName  |%ServiceDate       |
```


Add the ez-xl project into your solution, use the following snippet to use that file as a template and create multiple excel files 
from that template
```
public class User{
  public string UserName {get;set}
  public decimal Fee{get;set;}
  public DateTime ServiceDate{get;set;}
}
// ....
  using (var excelWriter = new ExcelWriter(@"c:\work\template.ods"))
  {
    for(var user in users)
    {
      User[] infoForCurrentUser = user.GetUserInfo();
      excelWriter.Write(string.Format("results-for-{0}.ods",user.UserId),
        new Tuple<string, IEnumerable<object>>("Sheet1",infoForCurrentUser),
        new Tuple<string, IEnumerable<object>>("Sheet2",infoForCurrentUser),
      );
    }
  }
```




