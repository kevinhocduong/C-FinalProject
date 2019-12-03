using SpotifyAPI;
using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WhatMoodAreYouIn.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult CreateDocument()
        {
            //Creates a PowerPoint instance
            IPresentation pptxDoc = Presentation.Create();

            //Adds a slide to the PowerPoint presentation
            ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);

            //Add a textbox to the slide
            IShape shape = slide.AddTextBox(10, 10, 500, 100);

            //Add a text to the textbox.
            shape.TextBody.AddParagraph("Hello World!!!");

            //Save the PowerPoint Presentation
            pptxDoc.Save("Sample.pptx", FormatType.Pptx, HttpContext.ApplicationInstance.Response);

            //Close the PowerPoint presentation
            pptxDoc.Close();

            return View();
        }
    }
}