asp.net c#從建立新題庫(NewQuestion)，選擇檔案到預覽(preview)，再選擇存檔(Submit)，尚未存檔時卻會直接跳回建立新題庫(NewQuestion)
並且這段報錯，說明Model.StationOptions 不存在為null，應如何處理
                    @Html.DropDownListFor(model => model.StationId,
                                          Model.StationOptions,
                                          "請選擇",
                                          new { @id = "station_id", onchange = "getStationId()" })

  

//後端Controller
public ActionResult Preview(HttpPostedFileBase txtFile, QuestionVM viewModel)
        {
            if (txtFile == null || txtFile.ContentLength == 0)
            {
                ModelState.AddModelError("", "請選擇 Excel 檔案！");
                return View("NewQuestion");
            }

            // 檢查副檔名
            var ext = Path.GetExtension(txtFile.FileName).ToLower();
            if (ext != ".xls")
            {
                ModelState.AddModelError("", "請上傳 .xls 格式的檔案！");
                return View("NewQuestion");
            }

            string fileName = Path.GetFileNameWithoutExtension(txtFile.FileName);
            string newFileName = fileName + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ext;
            string directoryPath = Server.MapPath("~/Content/Excel/");

            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            string path = Path.Combine(directoryPath, newFileName);
            txtFile.SaveAs(path);

            string htmlTable = ReadExcelToHtmlTable(txtFile.InputStream, viewModel);

            // 將 HTML 給 ViewModel 回傳給 Preview.cshtml 顯示
            var stationList = _TestingService.GetStationIdList().ToList();
            viewModel = new QuestionVM
            {
                StationList = stationList,
                StationOptions = stationList.Select(x => new SelectListItem
                {
                    Value = x,
                    Text = x
                }),
                Html = htmlTable,
                Filenames = newFileName
            };
            return View("Preview", viewModel);
        }

        public JsonResult GetCerItemIdList(string stationId)
        {
            var certItemList = _TestingService.GetCerItemIdList(stationId).ToList();
            var certItemOptions = certItemList.Select(x => new SelectListItem
            {
                Value = x,
                Text = x
            }).ToList();

            return Json(certItemOptions, JsonRequestBehavior.AllowGet);
        }

        public ActionResult NewQuestion()
        {
            var stationList = _TestingService.GetStationIdList().ToList();
            var viewModel = new QuestionVM
            {
                StationList = stationList,
                StationOptions = stationList.Select(x => new SelectListItem
                {
                    Value = x,
                    Text = x
                })
            };
            // 這裡可以預先帶入下拉選單等初始化資料
            return View("NewQuestion", viewModel);
        }

        public ActionResult Upload(HttpPostedFileBase txtFile, QuestionVM model)
        {
            string fileName = string.Empty;
            string newFileName = string.Empty;
            string path = string.Empty;
            if (model.Filenames == null)
            {
                if (txtFile == null || txtFile.ContentLength == 0)
                {
                    ModelState.AddModelError("", "請選擇 Excel 檔案！");
                    return View("NewQuestion");
                }

                // 檢查副檔名
                var ext = Path.GetExtension(txtFile.FileName).ToLower();
                if (ext != ".xls")
                {
                    ModelState.AddModelError("", "請上傳 .xls 格式的檔案！");
                    return View("NewQuestion");
                }

                fileName = Path.GetFileNameWithoutExtension(txtFile.FileName);
                newFileName = fileName + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ext;
                path = Path.Combine(Server.MapPath("~/Content/Excel/"), newFileName);
                txtFile.SaveAs(path);
                model.Filenames = newFileName;
            }
            // model.CerItemId, model.Filenames, model.StationId 等都會自動帶入
            // 你可以直接用 model.Filenames 來找 Excel 檔案處理
            string directoryPath = Server.MapPath("~/Content/Excel/");

            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            var excelPath = Path.Combine(Server.MapPath("~/Content/Excel/"), model.Filenames);
            var result = InsertExcelToOracle(excelPath, model.CerItemId);

            if (!string.IsNullOrEmpty(result))
            {
                TempData["Error"] = result;
                return RedirectToAction("Preview", model);
            }

            TempData["Success"] = "匯入成功";
            return RedirectToAction("Index");
        }


//前端View--------------------------------------------------------------------
@model MesTAManagementSystem_New.ViewModels.Training.Testing.QuestionVM
@using System.Web
@using System.Collections
@{

    ViewBag.Title = "NewQuestion(建立新題庫)";
    Layout = "~/Views/Shared/_LayoutMain.cshtml";
}
@section styles
{
    <link rel="stylesheet" href="~/Content/ses.css">
}
@{
    <script>
    function getStationId() {
        var stationId = $('#station_id').val();
        if (stationId) {
            $.ajax({
                url: '@Url.Action("GetCerItemIdList", "Question")',
                type: 'GET',
                data: { stationId: stationId },
                success: function (data) {
                    var $certItemSelect = $('#Cert_Item_Id');
                    $certItemSelect.empty(); // 先清空舊資料

                    if (data.length > 0) {
                        $certItemSelect.append($('<option>').val('').text('請選擇'));
                        $.each(data, function (index, item) {
                            $certItemSelect.append($('<option>').val(item.Value).text(item.Text));
                        });
                    } else {
                        $certItemSelect.append($('<option>').val('').text('無'));
                    }
                },
                error: function () {
                    alert('取得資料失敗');
                }
            });
        }
    }

    function addImg() {
        var tbl = document.getElementById('pictbl');
        var tr = tbl.insertRow(tbl.rows.length);
        var td = tr.insertCell(0);
        td.innerHTML = '圖檔上傳：<input type="file" name="Images" size="50" />';
    }

    function delImg() {
        var tbl = document.getElementById('pictbl');
        if (tbl.rows.length > 1) {
            tbl.deleteRow(tbl.rows.length - 1);
        } else {
            alert("不能再移除!");
        }
    }

    function disableImage() {
        var examSubject = document.getElementById("exam_subject").value;
        document.getElementById("pictbl").style.visibility = examSubject === "2" ? "hidden" : "visible";
    }

    function check(obj) {
        var fileName = document.getElementById("txtFile").value;
        if (!fileName) {
            alert("請輸入檔案名稱!");
            return false;
        }
        if (!fileName.endsWith(".xls")) {
            alert("檔案不是Excel檔!");
            return false;
        }
        if (document.getElementById("exam_subject").value === "1" &&
            document.getElementById("cert_item_id").value === "") {
            alert("請選取試題項目代碼!");
            return false;
        }
    }

    function setFormAction(actionName) {
        var baseUrl = '@Url.Content("~/")';
        var form = document.getElementById('uploadForm');
        form.action = baseUrl + 'Question/' + actionName;
    }

    </script>
}
<div id="NewQuestion-table-container">
    <form id="uploadForm" method="post" enctype="multipart/form-data" asp-controller="Question" onsubmit="return check(this)">
        <table width="600" cellpadding="5" id="NewQuestionTable">
            <tr>
                <td>
                    <img src="~/Content/images/title-blank.jpg" />
                    <span class="title">建立新題庫</span>
                </td>
            </tr>
            <tr><td><hr /></td></tr>
            <tr>
                <td>
                    <label>鑑定站別 :</label>
                    @Html.DropDownListFor(model => model.StationId,
                                          Model.StationOptions,
                                          "請選擇",
                                          new { @id = "station_id", onchange = "getStationId()" })

                    <label>筆試術科:</label>
                    <select id="exam_subject" name="ExamSubject" onchange="disableImage()">
                        <option value="1" selected="@(Model.ExamSubject == "1" ? "selected" : null)">筆試</option>
                        <option value="2" selected="@(Model.ExamSubject == "2" ? "selected" : null)">術科</option>
                    </select>

                    <label>試題項目代碼:</label>
                    <select id="Cert_Item_Id" name="CerItemId">
                        <option value="">請選擇</option>
                    </select>
                </td>
            </tr>

            <tr><td>Excel 檔案上傳：<input type="file" name="txtFile" id="txtFile" size="50" /></td></tr>
            <tr>
                <td>
                    <table id="pictbl">
                        <tr>
                            <td>
                                圖檔上傳：
                                <input type="file" name="Images" size="50" />
                                <input type="button" value="新增" onclick="addImg()" id="addPic" />
                                <input type="button" value="移除" onclick="delImg()" id="delPic" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input type="submit" value="預覽" name="action" onclick="setFormAction('Preview')" />
        <input type="submit" value="存檔" name="action" onclick="setFormAction('Upload')" />
    </form>
</div>
<script>
    if (document.getElementById("exam_subject").value === "2") {
        document.getElementById("img").disabled = true;
        document.getElementById("addPic").disabled = true;
        document.getElementById("delPic").disabled = true;
    }
</script>

//預覽
@model MesTAManagementSystem_New.ViewModels.Training.Testing.QuestionVM
@using System.Web
@using System.Collections
@{

    ViewBag.Title = "Preview(預覽)";
    var exam = Model.ExamSubject == "1" ? "筆試" : "術科";
    var showWeightAlert = Model.SCheckWeightFlag == "Y" && Model.ExamSubject == "2";
    Layout = "~/Views/Shared/_LayoutMain.cshtml";
}
@section styles
{
    <link rel="stylesheet" href="~/Content/ses.css">
}
@if (TempData["Success"] != null)
{
    <script>
        alert('@TempData["Success"]');
    </script>
}
@if (TempData["Error"] != null)
{
    <script>
        alert('@TempData["Error"]');
    </script>
}
<div id="NewQuestion-table-container">
    <form id="uploadForm" method="post" enctype="multipart/form-data" asp-controller="Question" onsubmit="return check(this)">
        @using (Html.BeginForm("Upload", "Question", FormMethod.Post, new { enctype = "multipart/form-data", name = "frm" }))
        {
            @Html.HiddenFor(m => m.Filenames)
            @Html.HiddenFor(m => m.StationId)
            @Html.HiddenFor(m => m.ExamSubject)
            @Html.HiddenFor(m => m.CerItemId)

            <table cellpadding="5">
                <tr>
                    <td><img src="~/Content/images/title-blank.jpg" /></td>
                    <td class="title">預覽</td>
                </tr>
            </table>
            <hr size="1" noshade="noshae" />

            <div>
                站別 : <b>@Model.StationId &nbsp; @exam</b>
                &nbsp; 試題項目代碼 : <b>@Model.CerItemId</b>
            </div>
            <hr size="1" noshade="noshae" />

            <div>
                @Html.Raw(Model.Html)
            </div>

            <input type="submit" name="action" value="存檔" />
            <input type="button" value="回上一頁" onclick="history.back()" />
        }

        @if (showWeightAlert)
        {
            <script>alert('術科權重欄位有小數點或%符號!');</script>
        }
    </form>
</div>

//ViewModel
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web;
using System.Web.Mvc;

namespace MesTAManagementSystem_New.ViewModels.Training.Testing
{
    public class QuestionVM
    {
        public List<string> StationList { get; set; }
        public List<string> CerItemList { get; set; }
        public IEnumerable<SelectListItem> StationOptions { get; set; }

        public IEnumerable<SelectListItem> CerItemOptions { get; set; }

        public string ExamSubject { get; set; }
        public List<string> VecItem { get; set; }
        public string QuestionType { get; set; }
        public string HtmlContent { get; set; }
        public string Filenames { get; set; }
        public string StationId { get; set; }
        public string CerItemId { get; set; }
        public string SCheckWeightFlag { get; set; }
        [Display(Name = "選項")]
        public bool IsChecked { get; set; }
        [Display(Name = "題型")]
        public string Type { get; set; }
        [Display(Name = "題號")]
        public string No { get; set; }
        [Display(Name = "子題")]
        public string Sub_No { get; set; }
        [Display(Name = "題目")]
        public string Subject { get; set; }
        [Display(Name = "檔名")]
        public string File_Name { get; set; }
        [Display(Name = "答案")]
        public string Answer { get; set; }
        [Display(Name = "負責人")]
        public string Sponsor { get; set; }
        [Display(Name = "可被選")]
        public string Delete_Flag { get; set; }
        public string RowId { get; set; }
        public List<string> deletedRowIds { get; set; }
        public string Html { get; set; } // ⬅ 用來顯示轉換後的表格

        public List<QuestionVM> EditQuestions { get; set; }

        public List<HttpPostedFileBase> Images { get; set; }
    }
}
