﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using ExcelAlias = Microsoft.Office.Interop.Excel;
namespace Excel
{
    public partial class Form1 : Form
    {
        RealEstateEntities context = new RealEstateEntities();
        List<Flat> Flats;
        ExcelAlias.Application xlApp;
        ExcelAlias.Workbook xlWB;
        ExcelAlias.Worksheet xlSheet;
        public Form1()
        {
            InitializeComponent();
            LoadData();
            CreateExcel(); 
        }

        private void CreateExcel()
        {
            throw new NotImplementedException();
        }

        private void LoadData()
        {
            Flats = context.Flat.ToList();
        }
    }
}
