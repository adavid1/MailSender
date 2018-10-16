using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Xml;
using System.Data.SqlClient;
using System.Drawing;
using OfficeOpenXml.Style;

namespace MailSender
{
    class ExcelReader
    {
        private int m_firstnameColumnID;
        private int m_lastNameColumnID;
        private int m_mailColumnID;
        private int m_regionColumnID;
        private int m_greetingColumnID;
        private int m_company;
        private int m_gender;

        private int m_ExpertsTitlesRow;

        public List<List<string>> m_expertsArray = new List<List<string>>();

        // Create an array of string with Lastname and email in each cells of blank status experts. 
        public void Init(string excelPath, List<string> codesRegion)
        {

            FileInfo existingFile = new FileInfo(excelPath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                bool flagFirstname = false;
                bool flagLastname = false;
                bool flagMail = false;
                bool flagRegion = false;
                bool flagGreeting = false;
                bool flagRowTitles = false;
                bool flagCompany = false;

                for (int row = worksheet.Dimension.Start.Row + 1; row < worksheet.Dimension.End.Row; row++)
                {
                    for (int columns = worksheet.Dimension.Start.Column; columns < worksheet.Dimension.End.Column + 1; columns++)
                    {
                        if (worksheet.Cells[row, columns].Text.Contains("Firstname") || worksheet.Cells[row, columns].Text.Contains("firstname"))
                        {
                            m_firstnameColumnID = columns;
                            flagFirstname = true;
                            flagRowTitles = true;
                        }

                        if (worksheet.Cells[row, columns].Text.Contains("Lastname") || worksheet.Cells[row, columns].Text.Contains("lastname"))
                        {
                            m_lastNameColumnID = columns;
                            flagLastname = true;
                        }

                        if (worksheet.Cells[row, columns].Text.Contains("Mail") || worksheet.Cells[row, columns].Text.Contains("mail"))
                        {
                            m_mailColumnID = columns;
                            flagMail = true;
                        }

                        if (worksheet.Cells[row, columns].Text.Contains("Company") || worksheet.Cells[row, columns].Text.Contains("company"))
                        {
                            m_company = columns;
                            flagCompany = true;
                        }

                        if (worksheet.Cells[row, columns].Text.Contains("Region") || worksheet.Cells[row, columns].Text.Contains("region"))
                        {
                            m_regionColumnID = columns;
                            flagRegion = true;
                        }

                        if (worksheet.Cells[row, columns].Text.Contains("Greeting") || worksheet.Cells[row, columns].Text.Contains("greeting"))
                        {
                            m_greetingColumnID = columns;
                            flagGreeting = true;
                        }

                        if (worksheet.Cells[row, columns].Text.Contains("Gender") || worksheet.Cells[row, columns].Text.Contains("gender"))
                        {
                            m_gender = columns;
                        }
                    }
                    if (flagRowTitles == true)
                    {
                        m_ExpertsTitlesRow = row;
                        row = worksheet.Dimension.End.Row;
                    }
                }
                

                if (flagLastname = false || flagMail == false || flagFirstname == false || flagRegion == false || flagGreeting == false || flagCompany== false)
                {
                    Console.WriteLine("columns are missing, try to rename or to add:");
                    Console.WriteLine("- Firstname");
                    Console.WriteLine("- Lastname");
                    Console.WriteLine("- Mail");
                    Console.WriteLine("- Company");
                    Console.WriteLine("- Region");
                    Console.WriteLine("- Greeting");
                    System.Threading.Thread.Sleep(8000);
                    Environment.Exit(1);
                }

                int expertID = 0;
                for (int row = m_ExpertsTitlesRow; row < worksheet.Dimension.End.Row + 1; row++)
                {
                    if ((worksheet.Cells[row, worksheet.Dimension.Start.Column].Value == null || (string)worksheet.Cells[row, worksheet.Dimension.Start.Column].Value == " "))
                    {
                        if (worksheet.Cells[row, m_lastNameColumnID].Value != null && worksheet.Cells[row, m_firstnameColumnID].Value != null)
                        {
                            if (codesRegion[0] == "4")
                            {
                                m_expertsArray.Add(new List<string>());

                                if (worksheet.Cells[row, m_firstnameColumnID].Value != null)
                                m_expertsArray[expertID].Add((string)worksheet.Cells[row, m_firstnameColumnID].Value);
                                m_expertsArray[expertID][0] = m_expertsArray[expertID][0].TrimEnd();
                                
                                m_expertsArray[expertID].Add((string)worksheet.Cells[row, m_lastNameColumnID].Value);
                                m_expertsArray[expertID][1] = m_expertsArray[expertID][1].TrimEnd();

                                m_expertsArray[expertID].Add((string)worksheet.Cells[row, m_mailColumnID].Value);

                                if (worksheet.Cells[row, m_regionColumnID].Value != null)
                                    m_expertsArray[expertID].Add(worksheet.Cells[row, m_regionColumnID].Value.ToString());
                                else
                                    m_expertsArray[expertID].Add("1");

                                if (worksheet.Cells[row, m_greetingColumnID].Value != null)
                                    m_expertsArray[expertID].Add(worksheet.Cells[row, m_greetingColumnID].Value.ToString());
                                else
                                    m_expertsArray[expertID].Add("1");

                                m_expertsArray[expertID].Add(row.ToString());

                                if (worksheet.Cells[row, m_company].Value != null)
                                {
                                    m_expertsArray[expertID].Add(worksheet.Cells[row, m_company].Value.ToString());
                                    m_expertsArray[expertID][6] = m_expertsArray[expertID][6].TrimEnd();
                                }
                                else
                                {
                                    m_expertsArray[expertID].Add("");
                                }

                                if (worksheet.Cells[row, m_gender].Value != null)
                                {
                                    m_expertsArray[expertID].Add(worksheet.Cells[row, m_gender].Value.ToString());
                                }
                                else
                                {
                                    m_expertsArray[expertID].Add("");
                                }

                                expertID++;

                            }
                            if (codesRegion.Count > 1)
                            {
                                if (worksheet.Cells[row, m_regionColumnID].Value != null)
                                {
                                    if (worksheet.Cells[row, m_regionColumnID].Value.ToString() == codesRegion[0] || worksheet.Cells[row, m_regionColumnID].Value.ToString() == codesRegion[1])
                                    {
                                        m_expertsArray.Add(new List<string>());

                                        m_expertsArray[expertID].Add((string)worksheet.Cells[row, m_firstnameColumnID].Value);
                                        m_expertsArray[expertID][0] = m_expertsArray[expertID][0].TrimEnd();

                                        m_expertsArray[expertID].Add((string)worksheet.Cells[row, m_lastNameColumnID].Value);
                                        m_expertsArray[expertID][1] = m_expertsArray[expertID][1].TrimEnd();

                                        m_expertsArray[expertID].Add((string)worksheet.Cells[row, m_mailColumnID].Value);

                                        if (worksheet.Cells[row, m_regionColumnID].Value != null)
                                            m_expertsArray[expertID].Add(worksheet.Cells[row, m_regionColumnID].Value.ToString());
                                        else
                                            m_expertsArray[expertID].Add("1");

                                        if (worksheet.Cells[row, m_greetingColumnID].Value != null)
                                            m_expertsArray[expertID].Add(worksheet.Cells[row, m_greetingColumnID].Value.ToString());
                                        else
                                            m_expertsArray[expertID].Add("1");

                                        m_expertsArray[expertID].Add(row.ToString());

                                        if (worksheet.Cells[row, m_company].Value != null)
                                        {
                                            m_expertsArray[expertID].Add(worksheet.Cells[row, m_company].Value.ToString());
                                            m_expertsArray[expertID][6] = m_expertsArray[expertID][6].TrimEnd();
                                        }
                                        else
                                        {
                                            m_expertsArray[expertID].Add("");
                                        }

                                        if (worksheet.Cells[row, m_gender].Value != null)
                                        {
                                            m_expertsArray[expertID].Add(worksheet.Cells[row, m_gender].Value.ToString());
                                        }
                                        else
                                        {
                                            m_expertsArray[expertID].Add("");
                                        }

                                        expertID++;
                                    }
                                }
                            }
                            if (codesRegion.Count == 1)
                            {
                                if (worksheet.Cells[row, m_regionColumnID].Value != null)
                                {
                                    if (worksheet.Cells[row, m_regionColumnID].Value.ToString() == codesRegion[0])
                                    {
                                        m_expertsArray.Add(new List<string>());

                                        m_expertsArray[expertID].Add((string)worksheet.Cells[row, m_firstnameColumnID].Value);
                                        m_expertsArray[expertID][0] = m_expertsArray[expertID][0].TrimEnd();

                                        m_expertsArray[expertID].Add((string)worksheet.Cells[row, m_lastNameColumnID].Value);
                                        m_expertsArray[expertID][1] = m_expertsArray[expertID][1].TrimEnd();

                                        m_expertsArray[expertID].Add((string)worksheet.Cells[row, m_mailColumnID].Value);

                                        if (worksheet.Cells[row, m_regionColumnID].Value != null)
                                            m_expertsArray[expertID].Add(worksheet.Cells[row, m_regionColumnID].Value.ToString());
                                        else
                                            m_expertsArray[expertID].Add("1");

                                        if (worksheet.Cells[row, m_greetingColumnID].Value != null)
                                            m_expertsArray[expertID].Add(worksheet.Cells[row, m_greetingColumnID].Value.ToString());
                                        else
                                            m_expertsArray[expertID].Add("1");

                                        m_expertsArray[expertID].Add(row.ToString());

                                        if (worksheet.Cells[row, m_company].Value != null)
                                        {
                                            m_expertsArray[expertID].Add(worksheet.Cells[row, m_company].Value.ToString());
                                            m_expertsArray[expertID][6] = m_expertsArray[expertID][6].TrimEnd();
                                        }
                                        else
                                        {
                                            m_expertsArray[expertID].Add("");
                                        }

                                        if (worksheet.Cells[row, m_gender].Value != null)
                                        {
                                            m_expertsArray[expertID].Add(worksheet.Cells[row, m_gender].Value.ToString());
                                        }
                                        else
                                        {
                                            m_expertsArray[expertID].Add("");
                                        }

                                        expertID++;
                                    }
                                }

                            }
                        }

                    }
                }
            }
        }

        public void WriteStatut(string excelPath, int row, int column)
        {
            FileInfo existingFile = new FileInfo(excelPath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                {
                    worksheet.SetValue(row, column, "emailed");   
                }
                package.Save();
            }
        }

        public List<List<string>> GetExpertsArray()
        {
            return m_expertsArray;
        }
    }
}
