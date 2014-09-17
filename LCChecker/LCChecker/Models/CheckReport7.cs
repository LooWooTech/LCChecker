﻿using LCChecker.Rules;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace LCChecker.Models
{
    public class CheckReport7 : CheckEngine,ICheck
    {


        public CheckReport7(string filePath)
        {
            GetMessage(filePath);
            var list = new List<IRowRule>();
            int count = Ship.Count();
            string[] IDS = new string[count];
            int i = 0;
            foreach (var item in Ship.Keys)
            {
                IDS[i] = item;
                i++;
            }
            list.Add(new CellRangeRowRule() { ColumnIndex = 3, Values = IDS });
            foreach (var item in Ship.Keys)
            {
                var rule1 = new StringEqual() { ColumnIndex = 3, Data = item };
                list.Add(new ConditionalRowRule()
                {
                    Condition = rule1,
                    Rule = new AndRule()
                    {
                        Rule1 = new AndRule()
                        {
                            Rule1 = new AndRule()
                            {
                                Rule1 = new StringEqual() { ColumnIndex = 1, Data = Ship[item].City },
                                Rule2 = new StringEqual() { ColumnIndex = 2, Data = Ship[item].County }
                            },
                            Rule2 = new AndRule()
                            {
                                Rule1 = new StringEqual() { ColumnIndex = 4, Data = Ship[item].Name },
                                Rule2 = new DoubleEqual() { ColumnIndex=5,data=Ship[item].AddArea}
                            }
                        },
                        Rule2 = new StringEqual() { ColumnIndex = 7, Data = Ship[item].Grade }
                    }
                });
            }
            list.Add(new CellRangeRowRule() { ColumnIndex = 8, Values = new[] { "是", "否" } });

            foreach (var item in list)
            {
                rules.Add(new RuleInfo() { Rule = item });
            }
        }


    }
}