using MVC.Test.Excel2Linq.Example.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVC.Test.Excel2Linq.Example.Controllers
{
    public class HomeController : Controller
    {
        /// <summary>
        /// Show the form to the user to submit the excel file
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// read the posted excel file and show the results
        /// </summary>
        /// <param name="myfile">excel file based on the template</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult List(HttpPostedFileBase myfile)
        {
            /// create a new ExcelToEntity instance based on the file uploaded by the user, 
            /// read it and transform it to a Listo of Person class
            var ListOfPersons = new ExcelToLinq.ExcelToEntity(myfile).Read<Person>();
            /// Return the List of Persons to the View to be displayed as HTML
            return View(ListOfPersons);
        }

    }
}