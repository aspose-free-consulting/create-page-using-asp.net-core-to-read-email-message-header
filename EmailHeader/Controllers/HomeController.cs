using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using EmailHeader.Models;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using Aspose.Email.Mapi;
using System.Data;

namespace EmailHeader.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {

            string SourceFilePath = "D:\\Input.msg";

            //Load the Message File
            SourceFilePath = @"/Users/mudassir/Documents/Email/Messages.msg";

            //Reading Email Msg file
            MapiMessage msg = MapiMessage.FromFile(SourceFilePath);

            //Reading keys
            var allKeys = msg.Headers.AllKeys;

            // Generate rows and cells.           
            int numRows = msg.Headers.Count() + 1;
            DataTable dt = new DataTable();
            dt.Columns.Add("HeaderName");
            dt.Columns.Add("Value");  

            //Iterating all Message headers and populating in table
            for (int rowNum = 1; rowNum < numRows; rowNum++)
            {
                DataRow dRow = dt.NewRow();
                dRow["HeaderName"] = allKeys[rowNum - 1].ToString();
                dRow["Value"] = msg.Headers[allKeys[rowNum - 1]].ToString();
                dt.Rows.Add(dRow);
            }

            return View(dt);
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
