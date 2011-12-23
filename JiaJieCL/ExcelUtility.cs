using System;
using System.Collections.Generic;
using System.Linq;
using System.Text; 

namespace JiaJieCL
{
    public class ExcelUtility
    {
        public Microsoft.Office.Interop.Excel.Application startExcel() 
        {
            Microsoft.Office.Interop.Excel.Application exc = new Microsoft.Office.Interop.Excel.Application();
            if (exc == null)
            {
                throw new Exception("Excel无法启动");
            }
            return exc;
        }

    }
}
