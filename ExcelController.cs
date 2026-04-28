using DocumentFormat.OpenXml.Office2010;
using DocumentFormat.OpenXml.Office2010.Excel;
using Excelform.ExcelDbContext;
using Excelform.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing;
using Microsoft.EntityFrameworkCore;
using StackExchange.Redis;
using System.Diagnostics.Contracts;
using System.Reflection.Metadata.Ecma335;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics.Metrics;
using Microsoft.AspNetCore.Mvc.Rendering;
using System.Collections.Generic;
using System.Net;

namespace Excelform.Controllers
{
    public class ExcelController : Controller
    {
        private readonly ApplicationDbContext applicationDbContext;

     
        public ExcelController(ApplicationDbContext applicationDbContext)
        {
            this.applicationDbContext = applicationDbContext;
        }
         
        [HttpGet] 
        public async Task<IActionResult> index(string searchString,int emp)
        {
            
                var employees = from e in applicationDbContext.emp_list
                                select e;

                if (!string.IsNullOrEmpty(searchString))
                {
                    employees = employees.Where(e =>
                        e.emp_id.ToString().Contains(searchString) ||
                        e.name.Contains(searchString) ||
                        e.designation.Contains(searchString) ||
                        e.department.Contains(searchString)
                     
                    );
                return View(await employees.OrderByDescending(x=>x.date_of_joining).ToListAsync());
            }
                else
                {
                    DateTime dateTime = DateTime.Now.AddDays(-45);
                    employees = employees.Where(e => e.date_of_joining >= dateTime);
                }

                return View(await employees.ToListAsync());
            }
            ////   var employee = await applicationDbContext.emp_list.ToListAsync();

            ////    var employees = from e in applicationDbContext.emp_list
            ////                    select e;
            ////    DateTime dateTime = DateTime.Now.AddDays(-45);
            ////    if (!string.IsNullOrEmpty(searchString))
            ////    {
            ////        employees = employees.Where(e =>
            ////            e.emp_id.ToString().Contains(searchString) ||
            ////            e.name.Contains(searchString) ||
            ////            e.designation.Contains(searchString) ||
            ////            e.department.Contains(searchString)
            ////        );

            ////        return View(employees.ToList());
            ////    }

            ////    //var filtereddata = applicationDbContext.emp_list.Where(x => EF.Functions.DateDiffDay(x.updated_on, DateTime.Now) <= 30).ToList();
            ////    // if (filtereddata != null)
            ////    //{
            ////    //    applicationDbContext.SaveChanges();
            ////    //return View(filtereddata);
            ////    // }

            ////    var filtereddata = applicationDbContext.emp_list.Where(x => x.date_of_joining >= dateTime).ToList();
            ////    return View(filtereddata);
            ////    //else
            ////    //{
            ////    //    applicationDbContext.SaveChanges();
            ////    //    return View(employee);
            ////    //}

            ////    // return View(employee);
            ////}

            [HttpGet]
        public IActionResult Add()
        {
            return View();

        }

        [HttpPost]
        public async Task<IActionResult> Add(Addview addemployeementrequest)
        {
            if (ModelState.IsValid)
            {
                var emplo = new Excel()
                {

                   // id = addemployeementrequest.id,
                    emp_id = addemployeementrequest.emp_id, //1
                    name = addemployeementrequest.name,//dinesh
                    designation = addemployeementrequest.designation,//techn
                    sub_department = addemployeementrequest.sub_department,//tech
                    gender = addemployeementrequest.gender,//m
                    //shift = addemployeementrequest.shift,
                    date_of_joining = addemployeementrequest.date_of_joining,//
                  //  date_of_resignation = addemployeementrequest.date_of_resignation,
                    location = addemployeementrequest.location,
                   // state = addemployeementrequest.state,
                   // status = addemployeementrequest.status,
                  //  company_classification = addemployeementrequest.company_classification,
                  //  employee_classification = addemployeementrequest.employee_classification,
                   created_on = addemployeementrequest.created_on,
                  // updated_on = addemployeementrequest.updated_on,
                    department = addemployeementrequest.department,
                  //  channel = addemployeementrequest.channel,
                  //  cluster = addemployeementrequest.cluster,
                  //  work_location = addemployeementrequest.work_location,
                  //  payroll_location = addemployeementrequest.payroll_location,
                    date_of_birth = addemployeementrequest.date_of_birth
                };
                await applicationDbContext.emp_list.AddAsync(emplo);
            };
            await applicationDbContext.SaveChangesAsync();

                return RedirectToAction("Add");
        }
        [HttpGet]

        public async Task<IActionResult> newview(int id)
        {
            //var model = applicationDbContext.emp_list.FirstOrDefaultAsync();
            var model = await applicationDbContext.emp_list.FirstOrDefaultAsync(x => x.id == id);

            if (model != null)
            {
                var viewmodel = new Excel()
                {
                    id = model.id,
                    emp_id = model.emp_id,
                    name = model.name,
                    designation = model.designation,
                    sub_department = model.sub_department,
                    gender = model.gender,
                    //shift = model.shift,
                    date_of_joining = (DateTime)model.date_of_joining,
                   // date_of_resignation = model.date_of_resignation,
                    //location = model.location,
                   // state = model.state,
                   // status = model.status,
                   // company_classification = model.company_classification,
                   // employee_classification = model.employee_classification,
                    created_on = model.created_on,
                   //updated_on = (DateTime)model.updated_on,
                    department = model.department,
                    //channel = model.channel,
                   // cluster = model.cluster,
                   // work_location = model.work_location,
                   // payroll_location = model.payroll_location,
                    date_of_birth = model.date_of_birth,
                };
                applicationDbContext.SaveChanges();

                return await Task.Run(() => View("newview", viewmodel));
            }
          //  return RedirectToAction("Index");
            return Json(model);
        }


        [HttpPost]

        public async Task<IActionResult> newview(Excel model)
        {
            var employee = await applicationDbContext.emp_list.FindAsync(model.id);

            if (employee != null)
            {
                employee.id = model.id;
                employee.emp_id = model.emp_id;
                employee.name = model.name;
                employee.designation = model.designation;
                employee.sub_department = model.sub_department;
                employee.gender = model.gender;
              //  employee.shift = model.shift;
               employee.date_of_joining = model.date_of_joining;
               // employee.date_of_resignation = model.date_of_resignation;
               // employee.location = model.location;
              //  employee.state = model.state;
              //  employee.status = model.status;
              //  employee.company_classification = model.company_classification;
              //  employee.employee_classification = model.employee_classification;
               //employee.created_on = model.created_on;
               // employee.updated_on = model.updated_on;
                employee.department = model.department;
                //employee.channel = model.channel;
              //  employee.cluster = model.cluster;
               // employee.work_location = model.work_location;
              //  employee.payroll_location = model.payroll_location;
                employee.date_of_birth = model.date_of_birth;
            }
         
           TempData["ConformationMessage"] = "Changes Saved Successfully.";
            applicationDbContext.Update(employee);
           
            applicationDbContext.SaveChanges();
            TempData["itemupdated"] = true;
           return RedirectToAction("Index");
          // return RedirectToAction("action");


        }
        public IActionResult action()
        {
            return View();

        }
        //public async Task<IActionResult> delete(int? id)
        //{
        //    Excel excel = applicationDbContext.emp_list.Find(id);
        //    applicationDbContext.emp_list.Remove(excel);
        //    applicationDbContext.SaveChanges();
        //    return RedirectToAction("Index");
        //}
        [HttpGet]
        public IActionResult search()
        {

            return View();
        }
            [HttpPost]
        public async Task<IActionResult> search(int emp)
        {
            if (emp != null)
            {   
                //var manager = await applicationDbContext.emp_list.FirstOrDefaultAsync();

                   DateTime dateTime = DateTime.Now.AddDays(-30);

                var permission = await applicationDbContext.emp_list
               .Where(x => x.emp_id == emp)
              .ToListAsync();


                if (emp == 4612)
                {
                    var employees = from e in applicationDbContext.emp_list
                                    select e;
                    DateTime date = DateTime.Now.AddDays(-45);
                    employees = employees.Where(e => e.date_of_joining >= date).OrderByDescending(x => x.date_of_joining);
                    return View("Index", employees);
                }
                if (emp == 4604)
                {
                    var employees = from e in applicationDbContext.emp_list
                                    select e;
                    DateTime date = DateTime.Now.AddDays(-45);
                    employees = employees.Where(e => e.date_of_joining >= date).OrderByDescending(x=>x.date_of_joining);
                    return View("Index", employees);
                }
                
               
            }
            return View("NO RECORDS FOUND");
        }

    }
}


//[HttpPost]

//public async Task<IActionResult> submit(Excel model)
//{

//    // var employee=applicationDbContext.emp_list.Where(x=>x.id==model.id).FirstOrDefault();
//    var employee = await applicationDbContext.emp_list.FindAsync(model);
//    if (employee != null)
//    {
//        employee.id = model.id;
//        employee.emp_id = model.emp_id;
//        employee.name = model.name;
//        employee.designation = model.designation;
//        employee.sub_department = model.sub_department;
//        employee.gender = model.gender;
//        employee.shift = model.shift;
//        employee.date_of_joining = model.date_of_joining;
//        employee.date_of_resignation = model.date_of_resignation;
//        employee.location = model.location;
//        employee.state = model.state;
//        employee.status = model.status;
//        employee.company_classification = model.company_classification;
//        employee.employee_classification = model.employee_classification;
//        employee.created_on = model.created_on;
//        employee.updated_on = model.updated_on;
//        employee.department = model.department;
//        employee.channel = model.channel;
//        employee.cluster = model.cluster;
//        employee.work_location = model.work_location;
//        employee.payroll_location = model.payroll_location;
//        employee.date_of_birth = model.date_of_birth;
//        applicationDbContext.SaveChanges();
//        //return RedirectToAction("Index");
//    }
//    return RedirectToAction("Index");
//}
//[HttpPost]
//public IActionResult edit(Excel emp)
//{
//    applicationDbContext.emp_list.Update(emp);
//    applicationDbContext.SaveChanges();
//   // return PartialView("edit", emp);
//    return RedirectToAction("Index");
//}
//[HttpPost]
//public async Task<IActionResult> index(ApplicationDbContext deletemod)
//{
//    var employee = await applicationDbContext.emp_list.FindAsync(deletemod);

//    applicationDbContext.emp_list.Remove(employee);
//    await applicationDbContext.SaveChangesAsync();
//    return RedirectToAction("Index");
//}
//public async Task<IActionResult> delete(int id)
//{
//    var employee = await applicationDbContext.emp_list.FirstOrDefaultAsync();

//    if (employee != null)
//    {
//        applicationDbContext.emp_list.Remove(employee);
//        await applicationDbContext.SaveChangesAsync();
//        return RedirectToAction("Index");
//    }
//    return RedirectToAction(nameof(Index));
//}

//[HttpGet]
//public ActionResult getemployee(int id)
//{
//    var employee = applicationDbContext.emp_list.Find(id);
//    return Json(employee.JsonRequestBehaviour);
//}
//[HttpPost]
//public ActionResult updateemployee(Excel updateemployee)
//{
//    if (ModelState.IsValid)
//    {
//        applicationDbContext.Entry(updateemployee).State = EntityState.Modified;
//        applicationDbContext.SaveChanges();
//    }
//    return Json(updateemployee);
//}