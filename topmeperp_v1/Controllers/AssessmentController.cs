using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace topmeperp.Controllers
{
    public class AssessmentController : Controller
    {
        // GET: Assessment
        public ActionResult Index()
        {
            return View();
        }

        // GET: Assessment
        public ActionResult AssessmentIndex()
        {
            Models.Assessment ass1 = new Models.Assessment();
            ass1.AssId=1;
            ass1.AssName = "水管";
            ass1.AssCost = 300;
                        
            return View(ass1);

        }
        public ActionResult AssessmentListIndex()
        {
            Models.Assessment ass1 = new Models.Assessment();
            ass1.AssId = 1;
            ass1.AssName = "水管";
            ass1.AssCost = 300;

            return View(ass1);
        }

    }
}