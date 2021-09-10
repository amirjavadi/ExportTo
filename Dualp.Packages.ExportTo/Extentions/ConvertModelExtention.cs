using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Dualp.Packages.ExportTo.Models;

namespace Dualp.Packages.ExportTo.Extentions
{
    public static class ConvertModelExtention
    {
        public static List<KeyValuesProperties> WriteTo<T>(this List<T> lst)
        {
            var keyvalues = new List<KeyValuesProperties>();


            var objType = typeof(T);


            //get name and display properties

            foreach (var propertyInfo in objType.GetProperties())
            {

                var name = propertyInfo.Name;
                var display = name;

                var diplayAttr = propertyInfo.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault();
                if (diplayAttr != null)
                    display = ((DisplayNameAttribute)diplayAttr).DisplayName;

                keyvalues.Add(new KeyValuesProperties()
                {
                    NameProperty = name,
                    DisplayProperty = display
                });
            }

            // get values 

            foreach (var obj in lst)
            {
                foreach (var propertyInfo in obj.GetType().GetProperties())
                {
                    var name = propertyInfo.Name;
                    var value = propertyInfo.GetValue(obj)?.ToString();

                    var findItem = keyvalues.FirstOrDefault(x => x.NameProperty == name);
                    if (findItem == null) continue;
                    findItem.Values ??= new List<string>();
                    findItem.Values.Add(value);
                }
            }


            return keyvalues;
        }
    }
}
