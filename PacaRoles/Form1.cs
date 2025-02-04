using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Data.Services.Client;
using PacaRoles.DataReference;
using Microsoft.Office;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace PacaRoles
{
    public partial class Form1 : Form
    {
        string filename;
        IdentityDataSource source;
        bool paca = false;
        int teamCount;
        Excel.Application excelPaca;
        Excel._Worksheet workSheet;

        public Form1()
        {
            InitializeComponent();            
        }

        private void btnGetFile_Click(object sender, EventArgs e)
        {
            dlgFilePick.Title = "Open File";
            dlgFilePick.InitialDirectory = @"C:\";
            if(dlgFilePick.ShowDialog() == DialogResult.OK)
            {
                filename = dlgFilePick.FileName;
                txtFilename.Text = filename;
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            if (filename == null)
            {
                MessageBox.Show("Error!  You have not selected a file.");
            }
            else
            {
                runProgram();
            }
            MessageBox.Show("Program completed.");
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            DateTime excelDate = DateTime.Now;
            string savePath = @FilePath + excelDate.ToString("MMddyy_HHmm") + ".xlsx";
            workSheet.SaveAs(savePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelPaca.Visible = true;
        }

        private void runProgram()
        {
            createExcel();
            teamCount = 1;
            var names = File.ReadAllLines(filename).ToList();

            foreach (var line in names)
            {
                teamCount++;
                paca = false;
                var parentName = line;                
                getTeamChildren(parentName, parentName);
                enterExcel(teamCount, paca, parentName);
            }

        }

        private void getTeamChildren(string parentName, string cookieCrumbs)
        {
            var parentId = getTeamID(parentName);

            if (parentId == null)
            {
                txtResults.AppendText("Parent Team " + parentName + " not found." + Environment.NewLine);
                return;
            }
            var childTeams = from t in source.Teams
                where t.ParentId == parentId
                select t;

            printParentChildren(cookieCrumbs, childTeams);
        }

        private void printParentChildren(string cookieCrumbs, IQueryable<Team> childTeams)
        {
            txtResults.AppendText(cookieCrumbs + Environment.NewLine);

            foreach (var team in childTeams)
            {
                if (team.Category != "Standard")
                {
                    txtResults.AppendText(" - " + team.DisplayName + Environment.NewLine);
                    var childId = getChildTeamId(team.DisplayName, team.ParentId);

                    if (childId == null)
                    {
                        txtResults.AppendText("No child team found for " + cookieCrumbs);
                        return;
                    }
                    var roles = from r in source.RoleAssignments
                                where r.ParentId == childId
                                select r;

                    printChildRoles(roles);
                }
            }
        }

        private void printChildRoles(IQueryable<RoleAssignment> childTeam)
        {
            foreach(var role in childTeam)
            {
                if(role.DisplayName == null)
                {
                    return;
                }
                if(role.DisplayName == "MSAuth.Physical Asset Change Authorizations")
                {
                    paca = true;
                    txtResults.AppendText("          Role - " + role.DisplayName + Environment.NewLine);
                }
            }
        }

        private string getTeamID(string teamName)
        {
            var rootTeam = from p in source.Teams
                           where p.DisplayName == teamName && p.BoundaryNodeType == "BRH.Application"
                           select p;
            var count = 0;
            string rootTeamId;

            foreach (var item in rootTeam)
                count++;

            if(count == 0)
            {
                return null;
            }
            try
            {
                rootTeamId = rootTeam.FirstOrDefault().Id;
            }
            catch
            {
                rootTeamId = null;
            }

            return rootTeamId;
        }

        private string getChildTeamId(string teamName, string parentID)
        {
            var childTeam = from t in source.Teams
                            where t.ParentId == parentID && t.DisplayName == teamName
                            select t;
                        
            var childTeamId = childTeam.FirstOrDefault().Id;

            return childTeamId;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            var IdentityURL = new Uri("");
            source = new DataReference.IdentityDataSource(IdentityURL)
            {
                MergeOption = MergeOption.OverwriteChanges,
                IgnoreMissingProperties = true,
                IgnoreResourceNotFoundException = true,
                Credentials = CredentialCache.DefaultNetworkCredentials
            };

            txtResults.AppendText("PACA Role Check:" + Environment.NewLine);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {           
            String txt = txtResults.Text;
            var date = DateTime.Now;
            TextWriter sw = new StreamWriter(Path.GetDirectoryName(filename) + @"\output_" + date.ToString("MMddyy_HHmm") + ".doc");
            sw.WriteLine(txt);
            sw.Close();

            if(File.Exists(Path.GetDirectoryName(filename) + @"\output_" + date.ToString("MMddyy_HHmm") + ".doc"))
            {
                MessageBox.Show("Information exported correctly");
            }
            else
            {
                MessageBox.Show("Export failed.");
            }
        }

        private void createExcel()
        {
            excelPaca = new Excel.Application();
            excelPaca.Workbooks.Add();

            workSheet = (Excel.Worksheet)excelPaca.ActiveSheet;

            workSheet.Cells[1, "A"] = "Property Dimension/Application";
            workSheet.Cells[1, "B"] = "PACA";

        }
        
        private void enterExcel(int row, bool role, string name)
        {
            string roleDes;

            if(role)
            {
                roleDes = "Y";
            }
            else
            {
                roleDes = "N";
            }

            workSheet.Cells[row, "A"] = name;
            workSheet.Cells[row, "B"] = roleDes;
        }
    }
}
