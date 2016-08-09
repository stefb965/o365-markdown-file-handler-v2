﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MarkdownFileHandler.Models
{
    public class FileHandlerModel
    {
        public FileHandlerModel(FileHandlerActivationParameters activationParameters, string fileContent, MvcHtmlString errorMessage)
        {
            this.ActivationParameters = activationParameters;
            if (null != FileContent)
                this.FileContent = new MvcHtmlString(fileContent);
            this.ErrorMessage = errorMessage;
        }

        public FileHandlerActivationParameters ActivationParameters { get; set; }
        public MvcHtmlString FileContent { get; set; }
        public MvcHtmlString ErrorMessage { get; set; }
    }
}