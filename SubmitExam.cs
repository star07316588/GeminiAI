        public ActionResult SubmitExam(OnlineExamVM model)
        {
            int NecsssaryIsCorrect = 0;
            int TureFalseIsCorrect = 0;
            int ChoiceIsCorrect = 0;
            int LinkIsCorrect = 0;
            int correct = 0;
            string Subject = string.Empty;
            string FileName = string.Empty;
            string UserAnswer = string.Empty;
            string Answer = string.Empty;

            if(model.ExamRegNo == null)
            {
                var ExamCerRegAndTime = _testingService.GetExamCerRegAndTime(model.EmpId, model.CerItemId);
                model.ExamRegNo = ExamCerRegAndTime.Select(m => m.ExamRegNo).FirstOrDefault();
            }

            // 必考題評分
            if (model.NecessaryQuestions != null)
            {
                for (int i = 0; i < model.NecessaryQuestions.Count; i++)
                {
                    Subject = model.NecessaryQuestions[i].Subject;
                    UserAnswer = model.NecessaryQuestions[i].UserAnswer;
                    Answer = model.NecessaryQuestions[i].Answer;
                    if (model.NecessaryQuestions[i].IsCorrect)
                    {
                        correct++;
                        NecsssaryIsCorrect++;
                    }
                }
            }

            // 是非題評分
            if (model.TrueFalseQuestions != null)
            {
                for (int i = 0; i < model.TrueFalseQuestions.Count; i++)
                {
                    Subject = model.TrueFalseQuestions[i].Subject;
                    UserAnswer = model.TrueFalseQuestions[i].UserAnswer;
                    Answer = model.TrueFalseQuestions[i].Answer;
                    if (model.TrueFalseQuestions[i].IsCorrect)
                    {
                        correct++;
                        TureFalseIsCorrect++;
                    }
                }
            }

            // 選擇題評分
            if (model.ChoiceQuestions != null)
            {
                for (int i = 0; i < model.ChoiceQuestions.Count; i++)
                {
                    Subject = model.ChoiceQuestions[i].Subject;
                    UserAnswer = model.ChoiceQuestions[i].UserAnswer;
                    Answer = model.ChoiceQuestions[i].Answer;
                    if (model.ChoiceQuestions[i].IsCorrect)
                    {
                        correct++;
                        ChoiceIsCorrect++;
                    }
                }
            }

            // 連連看評分
            if (model.LinkQuestions != null)
            {
                for (int i = 0; i < model.LinkQuestions.Count; i++)
                {
                    Subject = model.LinkQuestions[i].Subject;
                    UserAnswer = model.LinkQuestions[i].UserAnswer;
                    FileName = model.LinkQuestions[i].File_Name;
                    Answer = model.LinkQuestions[i].Answer;
                    if (model.LinkQuestions[i].IsCorrect)
                    {
                        correct++;
                        LinkIsCorrect++;
                    }
                }
            }

            model.TotalQuestions = (model.NecessaryQuestions != null ? model.NecessaryQuestions.Count : 0)
                                 + (model.TrueFalseQuestions != null ? model.TrueFalseQuestions.Count : 0)
                                 + (model.ChoiceQuestions != null ? model.ChoiceQuestions.Count : 0)
                                 + (model.LinkQuestions != null ? model.LinkQuestions.Count : 0);
            model.CorrectCount = correct;
            model.NecessaryScore = (int)((double)NecsssaryIsCorrect / model.TotalQuestions * 100);
            model.TrueFalseScore= (int)((double)TureFalseIsCorrect / model.TotalQuestions * 100);
            model.ChoiceScore = (int)((double)ChoiceIsCorrect / model.TotalQuestions * 100);
            model.LinkScore = (int)((double)LinkIsCorrect / model.TotalQuestions * 100);
            model.Score = (int)((double)correct / model.TotalQuestions * 100);

            model.ScoreType = "WR";
            bool success = _testingService.UpdateScore(model, userId);

            // 建立檔名與路徑
            string fileName = _testingService.GetPdfWritingFileName(model);
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = "ExamResult_" + model.EmpId + "_ " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
            }
            string folderPath = Server.MapPath("~/PDFResults/");
            string filePath = Path.Combine(folderPath, fileName);

            if(!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // 渲染 View 為 PDF
            var pdfView = new ViewAsPdf("ExamResultPdf", model)
            {
                PageSize = Rotativa.Options.Size.A4,
                CustomSwitches =
                    "--header-right \"General PD\nMacronix Proprietary\" " + 
                    "--header-font-size 9 " +
                    "--footer-center \"本文件內容涉及個人資料，僅供本公司業務範圍使用。\" " +
                    "--footer-font-size 8 " +
                    "--footer-line"
            };

            // 建立 PDF byte 並儲存於伺服器
            var pdfBytes = pdfView.BuildFile(ControllerContext);
            System.IO.File.WriteAllBytes(filePath, pdfBytes);

            return View("ExamResult", model); // 可建立 ExamResult.cshtml 顯示詳細評分與結果
        }

string path = Path.Combine(Server.MapPath("~/Content/Excel/"), newFileName);
