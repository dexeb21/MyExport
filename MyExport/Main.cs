using System;
using System.Collections.Generic;
using System.Text;
using FlexiCaptureScriptingObjects;
using System.Runtime.InteropServices;
using System.IO;
using System.Globalization;

namespace MyExportLibrary
{
    // Интерфейс компоненты экспорта, с которым можно работать из скрипта
    // При создании новой компоненты следует сгенерировать новый GUID
    [Guid("146CDAFE-1E57-4C6F-A5FB-86BB97745593")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IMyExport
    {
        [DispId(1)]
        //void ExportDocument(ref IExportDocument docRef, ref IExportTools exportTools);
        void ExportDocument(ref IExportDocument docRef, ref IExportTools exportTools, ref IExportImageSavingOptions imageOptions, ref string ExportFolder);
    }

    // Класс с реализацией функциональности компоненты экспорта
    // При создании новой компоненты следует сгенерировать новый GUID и 
    // задать свой ProgId
    [Guid("FA1B66CC-04A5-49BF-A0FA-35CB85EA7E50")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyExportLibrary.Export")]
    [ComVisible(true)]
    public class MyExport:IMyExport
    {
        public void ExportDocument(ref IExportDocument docRef, ref IExportTools exportTools, ref IExportImageSavingOptions imageOptions, ref string ExportFolder)
        {
            try
            {
                //string exportFolder = @"C:\MyExport";

                int val=0;
                Guid guid = Guid.NewGuid();
                Sections SectionsContainer = new Sections();
                exportSections(docRef.Children, ExportFolder, guid, ref val, ref SectionsContainer);
                exportImages(ref docRef, ref imageOptions, ExportFolder, ref SectionsContainer);
            }
            catch (Exception e)
            {
                docRef.Action.Succeeded = false;
                docRef.Action.ErrorMessage = e.ToString();
            }            
        }

///////////////////////////////////////////////////////////////////////////////////////////////////
        //private void exportImages(IExportDocument docRef, IExportTools exportTools, string exportFolder, ref Sections SectionsContainer)
        private void exportImages(ref IExportDocument docRef, ref IExportImageSavingOptions imageOptions, string exportFolder, ref Sections SectionsContainer)
        {
            /*IExportImageSavingOptions imageOptions = exportTools.NewImageSavingOptions();
            imageOptions.Format = "pdf";
            imageOptions.ColorType = "FullColor";
            imageOptions.Resolution = 300;*/
            //StreamWriter sw = File.CreateText(@"C:\MyExport\pages.txt");
            int k;
            /*try
            {*/
                for (int i = 0; i < SectionsContainer.Count; i++)
                {
                    //sw.WriteLine(SectionsContainer[i].FileName + " : " + SectionsContainer[i].StartInd);
                    //sw.WriteLine(i);
                    //sw.Flush();
                    foreach (IExportPage curPage in docRef.Pages)
                    {
                        curPage.ExcludedFromDocumentImage = true;
                    }
                    if (i == SectionsContainer.Count-1)
                    {
                        k = docRef.Pages.Count;
                    }
                    else
                    {
                        k = SectionsContainer[i + 1].StartInd;
                    }
                    for (int j = SectionsContainer[i].StartInd; j < k; j++)
                    {
                        docRef.Pages[j].ExcludedFromDocumentImage = false;
                    }

                    docRef.SaveAs(exportFolder + "\\" + SectionsContainer[i].FileName + ".pdf", imageOptions);
                }
            /*}
            //finally
            //{
              //  sw.Close();
            }*/
        }

///////////////////////////////////////////////////////////////////////////////////////////////////
        private void exportSections(IExportFields fields, string PathToExport, Guid guid, ref int section_num, ref Sections SectionsContainer)
        {
            foreach (IExportField curField in fields)
            {
                if (curField.Items != null & curField.Children == null)
                {
                    exportSections(curField.Items, PathToExport, guid, ref section_num, ref SectionsContainer);
                }

                if (curField.Items == null & curField.Children != null)
                {
                    section_num++;

                    String FileName = guid.ToString() + "_" + section_num.ToString("D3") + "_" + curField.Name;
                    SectionsContainer.Add(FileName, curField.Regions[0].PageIndex);

                    StreamWriter sw = File.CreateText(PathToExport +"\\"+ FileName + ".txt");
                    sw.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                    sw.WriteLine("<Document xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">");
                    sw.WriteLine("<Section Name=\"" + curField.Name + "\">");
                    exportAllFields(curField.Children, sw);
                    sw.WriteLine("</Section>");
                    sw.WriteLine("</Document>");
                    sw.Close();
                }
            }
        }

///////////////////////////////////////////////////////////////////////////////////////////////////
        private void exportAllFields(IExportFields fields, StreamWriter sw)
        {
            CultureInfo cultureUS, cultureRU;
            cultureUS = CultureInfo.CreateSpecificCulture("en-US");
            cultureRU = CultureInfo.CreateSpecificCulture("ru-RU");

            // Use standard numeric format specifiers.
                foreach (IExportField curField in fields)
                {
                    // сохраняем имя поля
                    if (curField.IsExportable)
                    {
                        sw.Write("   <{0}", curField.Name);

                        // сохраняем значение поля, если оно доступно
                        if (IsNullFieldValue(curField))
                        {
                            sw.WriteLine("></{0}>", curField.Name);
                        }
                        else
                        {                            
                            switch ((TExportFieldType)curField.Type)
                            {
                                case TExportFieldType.EFT_NumberField:
                                    if (!DBNull.Value.Equals(curField.Value))
                                    {
                                        sw.Write(">{0}", ((double)curField.Value).ToString(cultureUS));
                                    }
                                    else
                                    {
                                        sw.Write(">");
                                    }
                                    break;
                                case TExportFieldType.EFT_Checkmark:
                                    sw.Write(">{0}", curField.Value.ToString().ToUpper());
                                    break;
                                case TExportFieldType.EFT_DateTimeField:
                                    sw.Write(">{0}", curField.Text);
                                    break;
                                default:
                                    sw.Write(">{0}", curField.Text);
                                    break;
                            }
                            sw.WriteLine("</{0}>", curField.Name);
                        }
                    }
                    if (curField.Children != null)
                    {
                        // экспорт дочерних полей
                        exportAllFields(curField.Children, sw);
                    }
                    else if (curField.Items != null)
                    {
                        // экспорт экземпляров поля
                        exportAllFields(curField.Items, sw);
                    }
                }
        }
        
///////////////////////////////////////////////////////////////////////////////////////////////////
        // Проверяет значение поля на null
        // Если значение поля невалидно, то при любом обращении к 
        // нему(даже при проверке на null) может возникнуть исключение
        private bool IsNullFieldValue(IExportField field)
        {
            try
            {
                return (field.Value == null);
            }
            catch (Exception e)
            {
                return true;
            }
        }
///////////////////////////////////////////////////////////////////////////////////////////////////
    }
}