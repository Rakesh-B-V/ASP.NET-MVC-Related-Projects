namespace ReadExcelFile.Controllers
{
    using Dapper;
    using OfficeOpenXml;
    using ReadExcelFile.Models;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Data;
    using System.Data.SqlClient;
    using System.Linq;
    using System.Web;
    using System.Web.Mvc;
    public class HomeController : Controller
    {
        #region Global Declaration
        private static readonly String _dbConnection = ConfigurationManager.ConnectionStrings["dbConnection"].ConnectionString;
        #endregion Global Declaration

        #region Public Methods
        /// <summary>
        /// This method is used to Read the Excel and store data into Sql database
        /// </summary>
        /// <returns></returns>
        public ActionResult Upload()
        {
            var moviesList = new List<Movies>();
            if (Request != null)
            {
                HttpPostedFileBase file = Request.Files["UploadedFile"];
                if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                {
                    //String fileName = file.FileName;
                    //String fileContentType = file.ContentType;
                    //byte[] fileBytes = new byte[file.ContentLength];
                    //var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(file.InputStream))
                    {
                        var currentSheet = package.Workbook.Worksheets;
                        var worksheet = currentSheet.First();
                        var noOfCol = worksheet.Dimension.End.Column;
                        var noOfRow = worksheet.Dimension.End.Row;
                        for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
                        {
                            var movieDetails = new Movies()
                            {
                                MovieName = worksheet.Cells[rowIterator, 1].Value.ToString(),
                                Hero = worksheet.Cells[rowIterator, 2].Value.ToString(),
                                Director = worksheet.Cells[rowIterator, 3].Value.ToString()
                            };
                            moviesList.Add(movieDetails);
                        }
                    }
                    String insertQuery = String.Empty;
                    if (moviesList != null && moviesList.Any())
                    {
                        var lastItem = moviesList.LastOrDefault();
                        insertQuery = "INSERT INTO Movies (MovieName, Hero, Director) VALUES";
                        foreach (var item in moviesList)
                        {
                            if(item == lastItem)
                            {
                                insertQuery = insertQuery + "('" + item.MovieName + "','" + item.Hero + "','" + item.Director + "')";
                            }
                            else
                            {
                                insertQuery = insertQuery + "('" + item.MovieName + "','" + item.Hero + "','" + item.Director + "'),";
                            }
                        }
                    }
                    using (IDbConnection db = new SqlConnection(_dbConnection))
                    {
                        var insert = db.Query<dynamic>(insertQuery, commandType: CommandType.Text);
                    }
                }
            }
            
            return View("Upload",moviesList);
        }
        #endregion Public Methods
    }
}