using System;
using System.Collections.Generic;
using System.ComponentModel.Design.Serialization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Demo2m
{
    public class RowDataList : List<RowData>
    {

        public static List<string> ComparePCS(RowDataList list1, RowDataList list2)
        {
            return (from l1 in list1
                join l2 in list2 on l1.Brand equals l2.Brand
                where (l1.ComparedValue - l2.ComparedValue) > 1 || (l1.ComparedValue - l2.ComparedValue) < -1
                select l1.Brand + "/ " + l1.ComparedValue + " / "+ l2.ComparedValue + " / difference = " + (l1.ComparedValue - l2.ComparedValue)).ToList();
           
        }

        public static decimal CompareTotal(RowDataList list1, RowDataList list2)
        {
            var total1 = (from l1 in list1
                select l1.ComparedValue).Sum();

            var total2 = (from l2 in list2
                          select l2.ComparedValue).Sum();
            return total1 - total2;


        }
    }
}
