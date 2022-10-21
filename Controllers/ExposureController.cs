using Microsoft.AspNetCore.Mvc;
using ExposureTracker.Data;
using ExposureTracker.Models;
using System.Net.Http.Headers;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data;
using OfficeOpenXml;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Linq;
using ClosedXML.Excel;


namespace ExposureTracker.Controllers
{
    public class ExposureController :Controller
    {
        private readonly AppDbContext _db;
        IEnumerable<Insured> objInsuredList { get; set; }
        IEnumerable<TranslationTables> objTransTableList { get; set; }

        public ExposureController(AppDbContext db)
        {
            _db = db;
        }
        public IActionResult Index(string searchKey)
        {

            if(!string.IsNullOrEmpty(searchKey))
            {
                 searchKey = searchKey.Trim();
                objInsuredList = (from x in _db.dbLifeData where x.firstname.ToUpper().Contains(searchKey.ToUpper()) || (x.lastname.ToUpper().Contains(searchKey.ToUpper()) 
                || (x.policyno.Contains(searchKey.Trim()) || (x.baserider.Contains(searchKey.Trim()) || (x.plan.Contains(searchKey.Trim()) || (x.fullName.ToUpper().Contains(searchKey.ToUpper()) ||
                 (x.bordereauxfilename.ToUpper().Contains(searchKey.ToUpper()) || (x.cedingcompany.ToUpper().Contains(searchKey.ToUpper())))))))) select x).Take(50);
            }
            else
            {
                objInsuredList = _db.dbLifeData.OrderBy(x => x.cedingcompany).Take(100);
            }
            return View(objInsuredList);
        }

        public IActionResult Upload()
        {
            ViewBag.Message = "UPLOAD DATA TO DATABASE";
            return View();
        }


        //public ActionResult TestUpload(IFormFile file)
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Upload(IFormFile file, string selectDB)
        {
            int errorRowNo = 0;
            try
            {
                ViewBag.Message = "";
                string userName = Environment.UserName;
                if(file != null)
                {
                    if(selectDB == "SICS")
                    {
                        var list = new List<Insured>();
                        using(var stream = new MemoryStream())
                        {
                            await file.CopyToAsync(stream);
                            using(var package = new ExcelPackage(stream))
                            {
                                ExcelWorksheet worksheet = package.Workbook.Worksheets ["Sheet1"];
                                var rowcount = worksheet.Dimension.Rows;
                                for(int row = 2; row <= rowcount; row++)
                                {
                                    errorRowNo = row;
                                    list.Add(new Insured
                                    {
                                        identifier = worksheet.Cells [row, 1].Value.ToString().ToLower().Trim().ToUpper(),
                                        policyno = worksheet.Cells [row, 2].Value.ToString().Trim(),
                                        firstname = Convert.ToString(worksheet.Cells [row, 3].Value).Trim().ToUpper(),
                                        middlename = Convert.ToString(worksheet.Cells [row, 4].Value).Trim().ToUpper(),
                                        lastname = Convert.ToString(worksheet.Cells [row, 5].Value).Trim().ToUpper(),
                                        fullName = Convert.ToString(worksheet.Cells [row, 6].Value).Trim().ToUpper(),
                                        gender = Convert.ToString(worksheet.Cells [row, 7].Value).Trim().ToUpper(),
                                        clientid = Convert.ToString(worksheet.Cells [row, 8].Value).Trim().ToUpper(),
                                        dateofbirth = Convert.ToDateTime(worksheet.Cells [row, 9].Value).ToString("MM/dd/yyyy"),
                                        cedingcompany = Convert.ToString(worksheet.Cells [row, 10].Value).Trim().ToUpper(),
                                        cedantcode = Convert.ToString(worksheet.Cells [row, 11].Value).Trim(),
                                        typeofbusiness = Convert.ToString(worksheet.Cells [row, 12].Value).Trim().ToUpper(),
                                        bordereauxfilename = Convert.ToString(worksheet.Cells [row, 13].Value).Trim().ToUpper(),
                                        bordereauxyear = Convert.ToInt32(worksheet.Cells [row, 14].Value),
                                        soaperiod = Convert.ToString(worksheet.Cells [row, 15].Value),
                                        certificate = Convert.ToString(worksheet.Cells [row, 16].Value).Trim(),
                                        plan = Convert.ToString(worksheet.Cells [row, 17].Value).Trim().ToUpper(),
                                        benefittype = Convert.ToString(worksheet.Cells [row, 18].Value).Trim().ToUpper(),
                                        baserider = Convert.ToString(worksheet.Cells [row, 19].Value).Trim().ToUpper(),
                                        currency = Convert.ToString(worksheet.Cells [row, 20].Value).Trim(),
                                        planeffectivedate = Convert.ToDateTime(worksheet.Cells [row, 21].Value).ToString("MM/dd/yyyy"),
                                        sumassured = Decimal.Parse(worksheet.Cells [row, 22].Text.Trim()),
                                        //Convert.ToDecimal(worksheet.Cells [row, 22].Value),
                                        reinsurednetamountatrisk = Decimal.Parse(worksheet.Cells [row, 23].Text.Trim()),
                                        //Convert.ToDecimal(worksheet.Cells [row, 23].Value),
                                        mortalityrating = Convert.ToString(worksheet.Cells [row, 24].Value).Trim().ToUpper(),
                                        status = Convert.ToString(worksheet.Cells [row, 25].Value),
                                        dateuploaded = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss"),
                                        uploadedby = userName,
                                    });
                                        
                                }
                            }

                            list.ForEach(x =>
                            {
                                string strTransInsuranceProd = string.Empty;
                                string cedingComp = string.Empty;
                                var query = _db.dbLifeData.FirstOrDefault(y => y.identifier == x.identifier && y.policyno == x.policyno && y.plan == x.plan); //check current row in list if exists in the database
                                if(query != null)
                                {
                                    int listQuarter = fn_getQuarter(x.soaperiod);
                                    int queryQuarter = fn_getQuarter(query.soaperiod);
                                    //if(objInsuredList.Count() > 0)
                                    //{
                                    if(query.bordereauxyear <= x.bordereauxyear && listQuarter > queryQuarter && query.identifier == x.identifier && query.policyno == x.policyno && query.cedingcompany == x.cedingcompany && query.plan == x.plan)
                                    {

                                        if(!string.IsNullOrEmpty(query.benefittype))
                                        {
                                            query.identifier = x.identifier;
                                            query.policyno = x.policyno;
                                            query.firstname = x.firstname;
                                            query.middlename = x.middlename;
                                            query.lastname = x.lastname;
                                            query.fullName = x.fullName;
                                            query.gender = x.gender;
                                            query.clientid = x.clientid;
                                            query.dateofbirth = x.dateofbirth;
                                            query.cedingcompany = x.cedingcompany;
                                            query.typeofbusiness = x.typeofbusiness;
                                            query.bordereauxfilename = x.bordereauxfilename;
                                            query.bordereauxyear = x.bordereauxyear;
                                            query.soaperiod = x.soaperiod;
                                            query.certificate = x.certificate;
                                            query.plan = x.plan;
                                            var queryTransTable = _db.dbTranslationTable.FirstOrDefault(y => y.plan_code == x.plan && y.ceding_company == x.cedingcompany);
                                            query.cedantcode = x.cedantcode;
                                            query.baserider = fn_getBaseRider(queryTransTable.insured_prod);
                                            query.currency = x.currency;
                                            query.planeffectivedate = x.planeffectivedate;
                                            query.sumassured = x.sumassured;
                                            query.reinsurednetamountatrisk = x.reinsurednetamountatrisk;
                                            query.mortalityrating = x.mortalityrating;
                                            query.status = x.status;
                                            query.dateuploaded = x.dateuploaded;
                                            query.uploadedby = x.uploadedby;
                                            _db.Entry(query).State = EntityState.Modified;
                                        }
                                        else //benefit type column null in excel
                                        {
                                            query.identifier = x.identifier;
                                            query.policyno = x.policyno;
                                            query.firstname = x.firstname;
                                            query.middlename = x.middlename;
                                            query.lastname = x.lastname;
                                            query.fullName = x.fullName;
                                            query.gender = x.gender;
                                            query.clientid = x.clientid;
                                            query.dateofbirth = x.dateofbirth;
                                            query.cedingcompany = x.cedingcompany;
                                            query.typeofbusiness = x.typeofbusiness;
                                            query.bordereauxfilename = x.bordereauxfilename;
                                            query.bordereauxyear = x.bordereauxyear;
                                            query.soaperiod = x.soaperiod;
                                            query.certificate = x.certificate;
                                            query.plan = x.plan;
                                            var queryTransTable = _db.dbTranslationTable.FirstOrDefault(y => y.plan_code == x.plan && y.ceding_company == x.cedingcompany);
                                            query.cedantcode = queryTransTable.cedant_code;
                                            query.benefittype = queryTransTable.prod_description;// add prod description
                                            query.baserider = fn_getBaseRider(queryTransTable.insured_prod);
                                            query.currency = x.currency;
                                            query.planeffectivedate = x.planeffectivedate;
                                            query.sumassured = x.sumassured;
                                            query.reinsurednetamountatrisk = x.reinsurednetamountatrisk;
                                            query.mortalityrating = x.mortalityrating;
                                            query.status = x.status;
                                            query.dateuploaded = x.dateuploaded;
                                            query.uploadedby = x.uploadedby;
                                            _db.Entry(query).State = EntityState.Modified;
                                        }

                                    }// if bordereau year is less than the existing year in the database do nothhing
                                    #region exclude this logic for now
                                    else if(query.bordereauxyear < x.bordereauxyear)
                                    {
                                        query.identifier = x.identifier;
                                        query.policyno = x.policyno;
                                        query.firstname = x.firstname;
                                        query.middlename = x.middlename;
                                        query.lastname = x.lastname;
                                        query.fullName = x.fullName;
                                        query.gender = x.gender;
                                        query.clientid = x.clientid;
                                        query.dateofbirth = x.dateofbirth;
                                        query.cedingcompany = x.cedingcompany;
                                        var queryTransTable = _db.dbTranslationTable.FirstOrDefault(y => y.plan_code == x.plan && y.ceding_company == x.cedingcompany);
                                        query.cedantcode = queryTransTable.cedant_code;
                                        query.typeofbusiness = x.typeofbusiness;
                                        query.bordereauxfilename = x.bordereauxfilename;
                                        query.bordereauxyear = x.bordereauxyear;
                                        query.soaperiod = x.soaperiod;
                                        query.certificate = x.certificate;
                                        query.plan = x.plan;
                                        query.benefittype = queryTransTable.prod_description;// add prod description
                                        query.baserider = fn_getBaseRider(queryTransTable.insured_prod);
                                        query.currency = x.currency;
                                        query.planeffectivedate = x.planeffectivedate;
                                        query.sumassured = x.sumassured;
                                        query.reinsurednetamountatrisk = x.reinsurednetamountatrisk;
                                        query.mortalityrating = x.mortalityrating;
                                        query.status = x.status;
                                        _db.Entry(query).State = EntityState.Modified;
                                    }
                                    #endregion
                                    //}
                                }
                                else //current row in excel dont have record yet
                                {
                                    var newInsured = new Insured();
                                    if(query == null)
                                    {
                                        //if(x.benefittype == string.Empty || x.cedantcode == string.Empty)
                                        //{
                                            newInsured.identifier = x.identifier;
                                            newInsured.policyno = x.policyno;
                                            newInsured.firstname = x.firstname;
                                            newInsured.middlename = x.middlename;
                                            newInsured.lastname = x.lastname;
                                            newInsured.fullName = x.fullName;
                                            newInsured.gender = x.gender;
                                            newInsured.clientid = x.clientid;
                                            newInsured.dateofbirth = x.dateofbirth;
                                            newInsured.cedingcompany = x.cedingcompany;
                                            var queryTransTable = _db.dbTranslationTable.FirstOrDefault(y => y.plan_code == x.plan && y.ceding_company == x.cedingcompany);
                                            newInsured.cedantcode = queryTransTable.cedant_code;
                                            newInsured.typeofbusiness = x.typeofbusiness;
                                            newInsured.bordereauxfilename = x.bordereauxfilename;
                                            newInsured.bordereauxyear = x.bordereauxyear;
                                            newInsured.soaperiod = x.soaperiod;
                                            newInsured.certificate = x.certificate;
                                            newInsured.plan = queryTransTable.plan_code;
                                            newInsured.benefittype = queryTransTable.insured_prod;
                                            newInsured.baserider = fn_getBaseRider(queryTransTable.insured_prod);
                                            newInsured.currency = x.currency;
                                            newInsured.planeffectivedate = x.planeffectivedate;
                                            newInsured.sumassured = x.sumassured;
                                            newInsured.reinsurednetamountatrisk = x.reinsurednetamountatrisk;
                                            newInsured.mortalityrating = x.mortalityrating;
                                            newInsured.status = x.status;
                                            newInsured.dateuploaded = x.dateuploaded;
                                            newInsured.uploadedby = x.uploadedby;
                                        //}
                                    }
                                    _db.dbLifeData.Add(newInsured);
                                }
                            });
                        }
                        _db.SaveChanges();
                    }

                    //if the code reach here means everthing goes fine and excel data is imported into database
                    else if(selectDB == "TRANSLATION TABLE")
                    {

                        var list = new List<TranslationTables>();
                        using(var stream = new MemoryStream())
                        {
                            await file.CopyToAsync(stream);
                            using(var package = new ExcelPackage(stream))
                            {
                                ExcelWorksheet worksheet = package.Workbook.Worksheets ["Sheet1"];
                                var rowcount = worksheet.Dimension.Rows;
                                for(int row = 2; row <= rowcount; row++)
                                {

                                    list.Add(new TranslationTables
                                    {
                                        identifier = Convert.ToString(worksheet.Cells [row, 1].Value).Trim().ToUpper(),
                                        ceding_company = Convert.ToString(worksheet.Cells [row, 2].Value).Trim().ToUpper(),
                                        cedant_code = Convert.ToString(worksheet.Cells [row, 3].Value).Trim().ToUpper(),
                                        plan_code = Convert.ToString(worksheet.Cells [row, 4].Value).Trim().ToUpper(),
                                        benefit_cover = Convert.ToString(worksheet.Cells [row, 5].Value).Trim().ToUpper(),
                                        insured_prod = Convert.ToString(worksheet.Cells [row, 6].Value).Trim().ToUpper(),
                                        prod_description = Convert.ToString(worksheet.Cells [row, 7].Value).Trim().ToUpper(),
                                    });

                                }
                            }
                            foreach(var x in list)
                            {
                                var query = _db.dbTranslationTable.FirstOrDefault(y => y.identifier == x.identifier && y.ceding_company == x.ceding_company);
                                if(query != null)
                                {
                                    query.identifier = x.identifier;
                                    query.ceding_company = x.ceding_company;
                                    query.plan_code = x.plan_code;
                                    query.cedant_code = x.cedant_code;
                                    query.benefit_cover = x.benefit_cover;
                                    query.insured_prod = x.insured_prod;
                                    query.prod_description = x.prod_description;
                                    query.base_rider = fn_getBaseRider(x.insured_prod);
                                    _db.Entry(query).State = EntityState.Modified;

                                }
                                else
                                {
                                    var newRecord = new TranslationTables();
                                    if(string.IsNullOrEmpty(x.base_rider))
                                    {
                                        newRecord.identifier = x.identifier;
                                        newRecord.plan_code = x.plan_code;
                                        newRecord.base_rider = fn_getBaseRider(x.insured_prod);
                                        newRecord.insured_prod = x.insured_prod;
                                        newRecord.ceding_company = x.ceding_company;
                                        newRecord.prod_description = x.prod_description;
                                        newRecord.benefit_cover = x.benefit_cover;
                                        newRecord.cedant_code = x.cedant_code;
                                    }
                                    _db.dbTranslationTable.Add(newRecord);
                                    //_db.AddRange(listTable);
                                }

                            }
                            _db.SaveChanges();
                        }
                    }

                    #region BasicRider uploading
                    //else if(selectedDB == "BasicRider")
                    //{

                    //    var list = new List<BasicRiderProd>();
                    //    using(var stream = new MemoryStream())
                    //    {
                    //        await DBTemplate.CopyToAsync(stream);
                    //        using(var package = new ExcelPackage(stream))
                    //        {
                    //            ExcelWorksheet worksheet = package.Workbook.Worksheets ["Sheet1"];
                    //            var rowcount = worksheet.Dimension.Rows;
                    //            for(int row = 2; row <= rowcount; row++)
                    //            {

                    //                list.Add(new BasicRiderProd
                    //                {
                    //                    insuredprod_basic = Convert.ToString(worksheet.Cells [row, 1].Value).Trim().ToUpper(),
                    //                    insuredprod_rider = Convert.ToString(worksheet.Cells [row, 2].Value).Trim().ToUpper(),

                    //                });

                    //            }
                    //        }


                    //        foreach(var x in list)
                    //        {
                    //            var query = _db.dbBasicRider.FirstOrDefault(y => y.insuredprod_basic == x.insuredprod_basic);
                    //            if(query != null)
                    //            {
                    //                query.insuredprod_basic = x.insuredprod_basic;
                    //                query.insuredprod_rider = x.insuredprod_rider;
                    //                _db.Entry(query).State = EntityState.Modified;

                    //            }
                    //            else
                    //            {

                    //                 var newRecord = new BasicRiderProd();
                    //                 newRecord.insuredprod_basic = x.insuredprod_basic;
                    //                 newRecord.insuredprod_rider = x.insuredprod_rider;
                    //                _db.dbBasicRider.Add(newRecord);
                    //                //_db.AddRange(listTable);
                    //            }

                    //        }
                    //        _db.SaveChanges();
                    //    }
                    //}
                    #endregion
                    else
                    {
                        ViewBag.Message = "SELECT A DATABASE";
                        return View("Upload");
                    }
                    ViewBag.Message = selectDB.ToUpper() + " has been Uploaded to Database";
                    return View("Upload");
                }

                else
                {
                    ViewBag.Message = "UPLOAD A FILE AND SELECT A DATABASE";
                    return View("Upload");
                }

            }
            catch(Exception ex)
            {
                errorRowNo = errorRowNo - 1;
                ViewBag.Message = "Ceding Company and Plan code should match in the translation table, check row no : " + errorRowNo + " in your input file";
                return View("Upload");
            }

        }

        #region Insured Prod Function
        //public List<Insured> fn_checkInsuredProd(List<>, string value_BenefitType)
        //{

        //    switch (value_BenefitType)
        //    {
        //        case "123":
        //        var newADB_IND = new ADB_IND();
        //        var lstADB_IND = new List<ADB_IND>();
        //        foreach(var item in listRiders)
        //        {
        //            newADB_IND.rider = item.baserider;
        //            newADB_IND.insuredprod = item.benefittype;
        //            newADB_IND.totalSumAssured += item.sumassured;
        //            newADB_IND.totalNAR += item.reinsurednetamountatrisk;
        //            lstADB_IND.Add(newADB_IND);
        //            return lstADB_IND;
        //        }
        //        break;



        //        default:
        //        var newACCD_DRGRP = new ADB_IND();
        //        var lstACCD_DRGRP = new List<ADB_IND>();
        //        lstACCD_DRGRP.Add(newACCD_DRGRP);
        //        return lstACCD_DRGRP;
        //        break;
        //    }

        //    return null;

        //}
        #endregion


        public IActionResult ViewAccumulation()
        {
            return View();
        }


        public IActionResult ViewPolicies(string Identifier)
        {
            objInsuredList = (from obj in _db.dbLifeData
                              where obj.identifier.Contains(Identifier)
                              select obj).ToList();


            string strFullName = string.Empty;
            string strIdentifier = string.Empty;

            foreach(var item in objInsuredList)
            {
                strIdentifier = item.identifier;
                strFullName = item.fullName.ToUpper();
                if(strFullName != string.Empty)
                {
                    break;
                }
                else
                {
                    continue;
                }
            }
            TempData ["Identifier"] = Identifier;
            ViewBag.FullName = strFullName;
            return View(objInsuredList);
        }

        public IActionResult EditSession(int Id)
        {
            var objInsured = _db.dbLifeData.Find(Id);
            objInsured.dateuploaded = DateTime.Now.ToString("MM/dd/yyy");
            return PartialView("_partialViewEdit", objInsured);
        }


        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Edit(Insured objInsuredList)
        {
            _db.dbLifeData.Update(objInsuredList);
            _db.SaveChanges();
            return RedirectToAction("Index");

        }

        public int fn_getQuarter(string value)
        {
            string quarter = string.Empty;
            string quarter_ = string.Empty;
            int quarterNo = 0;
            var number = Regex.Matches(value, @"\d+");
            foreach(var no in number)
            {
                quarter += no;
            }
            quarter_ = quarter;
            quarterNo = int.Parse(quarter_);
            return quarterNo;
        }
        public string fn_getBaseRider(string valueInsuranceProd)
        {
            string [] InsuranceProd = { "VARIABLELIFE-RE", "TRADITIONALLIFE", "TERMLIFE-GRP", "CREDITLIFE-GRP" };

            foreach(var item in InsuranceProd)
            {
                if(item.ToUpper() == valueInsuranceProd.ToUpper())
                {
                    return "BASIC";
                    break;

                }
                else
                {
                    continue;

                }
            }
            return "RIDER";

        }

        public IActionResult ViewDetails(string Identifier)
        {
            TableBasicRider tableBasicRider = new TableBasicRider();


            var Account = _db.dbLifeData.Where(y => y.identifier == Identifier);
            var selectBasics = _db.dbLifeData.Where(y => y.identifier == Identifier && y.baserider == "BASIC");
            var selectRiders = _db.dbLifeData.Where(y => y.identifier == Identifier && y.baserider == "RIDER");

            int intPolicyNo = Account.Count();
            var userDetails = Account.FirstOrDefault(x => x.identifier == Identifier);
            string strFullname = userDetails.fullName;
            string strDOB = userDetails.dateofbirth;

            #region Model List
            var newBasic = new BASIC();
            var newACCD_DRGRP = new ACCD_DRGRP();
            var newACCDDIS_DGRP = new ACCDDIS_DGRP();
            var newACCDEEATHIND = new ACCDEEATHIND();
            var newACCDISBENIND = new ACCDISBENIND();
            var newACCIDENTALDEATH = new ACCIDENTALDEATH();
            var newACCIDNTDTHDISAB = new ACCIDNTDTHDISAB();
            var newADDGRP = new ADDGRP();
            var newADDIND = new ADDIND();
            var newAADBDIND = new ADBDIND();
            var newADBI = new ADBI();
            var newADBIND = new ADBIND();
            var newADBRGRP = new ADBRGRP();
            var newADDDIND = new ADDDIND();
            var newADPGRP = new ADPGRP();
            var newADPIND = new ADPIND();
            var newBBGRP = new BBGRP();
            var newCIENDSRIND = new CIENDSRIND();
            var newCIESIND = new CIESIND();
            var newCIRACIND = new CIRACIND();
            var newCIRNAIND = new CIRNAIND();
            var newCRITICALILLNESS = new CRITICALILLNESS();
            var newDHIACCIND = new DHIACCIND();
            var newDHIBACGRP = new DHIBACGRP();
            var newDHIBALLGRP = new DHIBALLGRP();
            var newDHIBALLIND = new DHIBALLIND();
            var newDHIBILIND = new DHIBILIND();
            var newDOUBLEINDIND = new DOUBLEINDIND();
            var newMCFRAIND = new MCFRAIND();
            var newMCFRININD = new MCFRININD();
            var newMEDICALREIIND = new MEDICALREIIND();
            var newMEDICALREIMBURS = new MEDICALREIMBURS();
            var newMEDICALREIMGRP = new MEDICALREIMGRP();
            var newMEDIINSIND = new MEDIINSIND();
            var newMORTGAGEREDEMPT = new MORTGAGEREDEMPT();
            var newMRBACCIND = new MRBACCIND();
            var newMRPGRP = new MRPGRP();
            var newMURDERASSAULT = new MURDERASSAULT();
            var newPTDISINCOIND = new PTDISINCOIND();
            var newRENEWALPERSONAL = new RENEWALPERSONAL();
            var newRPAR = new RPAR();
            var newSACIENDSTAPIND = new SACIENDSTAPIND();
            var newSACIESPIND = new SACIESPIND();
            var newSADBIND = new SADBIND();
            var newSPLADBIND = new SPLADBIND();
            var newSTANDALONECRITI = new STANDALONECRITI();
            var newSTANDALONEENH = new STANDALONEENH();
            var newTPDINCOMEGRP = new TPDINCOMEGRP();
            var newTPDLSGRP = new TPDLSGRP();
            var newTPDLSIND = new TPDLSIND();
            var newTTDISINCOIND = new TTDISINCOIND();
            var newTTDISLSGRP = new TTDISLSGRP();
            var newTTDLSIND = new TTDLSIND();
            var newTEMRIDISNIND = new TEMRIDISNIND();
            var newTERMRIDERPAYOR = new TERMRIDERPAYOR();
            var newTIR = new TIR();
            var newVARLIFEGU = new VARLIFEGU();
            var newWOPDDPIND = new WOPDDPIND();
            var newWOPDDIIND = new WOPDDIIND();
            var newWOPDIIND = new WOPDIIND();
            var newWOPDIIND_ = new WOPDIIND_();
            var newWOPDOPIND = new WOPDOPIND();
            var newWOPDPIND = new WOPDPIND();


            var lstBasic = new List<BASIC>();
            var lstACCD_DRGRP = new List<ACCD_DRGRP>();
            var lstACCDDIS_DGRP = new List<ACCDDIS_DGRP>();
            var lstACCDEEATHIND = new List<ACCDEEATHIND>();
            var lstACCDISBENIND = new List<ACCDISBENIND>();
            var lstACCIDENTALDEATH = new List<ACCIDENTALDEATH>();
            var lstACCIDNTDTHDISAB = new List<ACCIDNTDTHDISAB>();
            var lstADDGRP = new List<ADDGRP>();
            var lstADDIND = new List<ADDIND>();
            var lstADBDIND = new List<ADBDIND>();
            var lstADBI = new List<ADBI>();
            var lstADBIND = new List<ADBIND>();
            var lstADBRGRP = new List<ADBRGRP>();
            var lstADDDIND = new List<ADDDIND>();
            var lstADPGRP = new List<ADPGRP>();
            var lstADPIND = new List<ADPIND>();
            var lstBBGRP = new List<BBGRP>();
            var lstCIENDSRIND = new List<CIENDSRIND>();
            var lstCCIESIND = new List<CIESIND>();
            var lstCIRACIND = new List<CIRACIND>();
            var lstCIRNAIND = new List<CIRNAIND>();
            var lstCRITICALILLNESS = new List<CRITICALILLNESS>();
            var lstDHIACCIND = new List<DHIACCIND>();
            var lstDHIBACGRP = new List<DHIBACGRP>();
            var lstDHIBALLGRP = new List<DHIBALLGRP>();
            var lstDHIBALLIND = new List<DHIBALLIND>();
            var lstDHIBILIND = new List<DHIBILIND>();
            var lstDOUBLEINDIND = new List<DOUBLEINDIND>();
            var lstMCFRAIND = new List<MCFRAIND>();
            var lstMCFRININD = new List<MCFRININD>();
            var lstMEDICALREIIND = new List<MEDICALREIIND>();
            var lstMEDICALREIMBURS = new List<MEDICALREIMBURS>();
            var lstMEDICALREIMGRP = new List<MEDICALREIMGRP>();
            var lstMEDIINSIND = new List<MEDIINSIND>();
            var lstMORTGAGEREDEMPT = new List<MORTGAGEREDEMPT>();
            var lstMRBACCIND = new List<MRBACCIND>();
            var lstMRPGRP = new List<MRPGRP>();
            var lstMURDERASSAULT = new List<MURDERASSAULT>();
            var lstPTDISINCOIND = new List<PTDISINCOIND>();
            var lstRENEWALPERSONAL = new List<RENEWALPERSONAL>();
            var lstRPAR = new List<RPAR>();
            var lstSACIENDSTAPIND = new List<SACIENDSTAPIND>();
            var lstSACIESPIND = new List<SACIESPIND>();
            var lstSADBIND = new List<SADBIND>();
            var lstSPLADBIND = new List<SPLADBIND>();
            var lstSTANDALONECRITI = new List<STANDALONECRITI>();
            var lstSTANDALONEENH = new List<STANDALONEENH>();
            var lstTPDINCOMEGRP = new List<TPDINCOMEGRP>();
            var lstTPDLSGRP = new List<TPDLSGRP>();
            var lstTPDLSIND = new List<TPDLSIND>();
            var lstTTDISINCOIND = new List<TTDISINCOIND>();
            var lstTTDISLSGRP = new List<TTDISLSGRP>();
            var lstTTDLSIND = new List<TTDLSIND>();
            var lstTEMRIDISNIND = new List<TEMRIDISNIND>();
            var lstTERMRIDERPAYOR = new List<TERMRIDERPAYOR>();
            var lstTIR = new List<TIR>();
            var lstVARLIFEGU = new List<VARLIFEGU>();
            var lstWOPDDPIND = new List<WOPDDPIND>();
            var lstWOPDDIIND = new List<WOPDDIIND>();
            var lstWOPDIIND = new List<WOPDIIND>();
            var lstWOPDIIND_ = new List<WOPDIIND_>();
            var lstWOPDOPIND = new List<WOPDOPIND>();
            var lstWOPDPIND = new List<WOPDPIND>();
            #endregion
            #region Loop  to all BASICS
            foreach(var item in selectBasics)
            {
                newBasic.basic = item.baserider;
                newBasic.insuredprod = "TOTAL BASIC";
                newBasic.totalSumAssured += item.sumassured;
                newBasic.totalNAR += item.reinsurednetamountatrisk;
            }
            lstBasic.Add(newBasic);
            tableBasicRider.BASICS = lstBasic;
            #endregion

            #region Loop to all RIDERS
            foreach(var item in selectRiders)
            {
                if(item.benefittype == "ACCD&DRGRP")
                {
                    newACCD_DRGRP.rider = item.baserider;
                    newACCD_DRGRP.insuredprod = item.benefittype;
                    newACCD_DRGRP.totalSumAssured += item.sumassured;
                    newACCD_DRGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstACCD_DRGRP.Add(newACCD_DRGRP);
                    tableBasicRider.ACCD_DRGRP = lstACCD_DRGRP;

                }
                else if(item.benefittype == "ACCDDIS_DGRP")
                {
                    newACCDDIS_DGRP.rider = item.baserider;
                    newACCDDIS_DGRP.insuredprod = item.benefittype;
                    newACCDDIS_DGRP.totalSumAssured += item.sumassured;
                    newACCDDIS_DGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstACCDDIS_DGRP.Add(newACCDDIS_DGRP);
                    tableBasicRider.ACCDDIS_DGRP = lstACCDDIS_DGRP;
                }
                else if(item.benefittype == "ACCDEEATHIND")
                {
                    newACCDEEATHIND.rider = item.baserider;
                    newACCDEEATHIND.insuredprod = item.benefittype;
                    newACCDEEATHIND.totalSumAssured += item.sumassured;
                    newACCDEEATHIND.totalNAR += item.reinsurednetamountatrisk;
                    lstACCDEEATHIND.Add(newACCDEEATHIND);
                    tableBasicRider.ACCDEEATHIND = lstACCDEEATHIND;
                }
                else if(item.benefittype == "ACCDISBENIND")
                {
                    newACCDEEATHIND.rider = item.baserider;
                    newACCDEEATHIND.insuredprod = item.benefittype;
                    newACCDEEATHIND.totalSumAssured += item.sumassured;
                    newACCDEEATHIND.totalNAR += item.reinsurednetamountatrisk;
                    lstACCDISBENIND.Add(newACCDISBENIND);
                    tableBasicRider.ACCDISBENIND = lstACCDISBENIND;
                }
                else if(item.benefittype == "ACCIDENTALDEATH")
                {
                    newACCIDENTALDEATH.rider = item.baserider;
                    newACCIDENTALDEATH.insuredprod = item.benefittype;
                    newACCIDENTALDEATH.totalSumAssured += item.sumassured;
                    newACCIDENTALDEATH.totalNAR += item.reinsurednetamountatrisk;
                    lstACCIDENTALDEATH.Add(newACCIDENTALDEATH);
                    tableBasicRider.ACCIDENTALDEATH = lstACCIDENTALDEATH;
                }
                else if(item.benefittype == "ACCIDNTDTHDISAB")
                {
                    newACCIDNTDTHDISAB.rider = item.baserider;
                    newACCIDNTDTHDISAB.insuredprod = item.benefittype;
                    newACCIDNTDTHDISAB.totalSumAssured += item.sumassured;
                    newACCIDNTDTHDISAB.totalNAR += item.reinsurednetamountatrisk;
                    lstACCIDNTDTHDISAB.Add(newACCIDNTDTHDISAB);
                    tableBasicRider.ACCIDNTDTHDISAB = lstACCIDNTDTHDISAB;
                }
                else if(item.benefittype == "AD&D-GRP")
                {
                    newADDGRP.rider = item.baserider;
                    newADDGRP.insuredprod = item.benefittype;
                    newADDGRP.totalSumAssured += item.sumassured;
                    newADDGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstADDGRP.Add(newADDGRP);
                    tableBasicRider.ADDGRP = lstADDGRP;
                }
                else if(item.benefittype == "AD&D-IND")
                {
                    newADDIND.rider = item.baserider;
                    newADDIND.insuredprod = item.benefittype;
                    newADDIND.totalSumAssured += item.sumassured;
                    newADDIND.totalNAR += item.reinsurednetamountatrisk;
                    lstADDIND.Add(newADDIND);
                    tableBasicRider.ADDIND = lstADDIND;
                }
                else if(item.benefittype == "ADB&D-IND")
                {
                    newAADBDIND.rider = item.baserider;
                    newAADBDIND.insuredprod = item.benefittype;
                    newAADBDIND.totalSumAssured += item.sumassured;
                    newAADBDIND.totalNAR += item.reinsurednetamountatrisk;
                    lstADBDIND.Add(newAADBDIND);
                    tableBasicRider.ADBDIND = lstADBDIND;
                }
                else if(item.benefittype == "ADB-I")
                {
                    newADBI.rider = item.baserider;
                    newADBI.insuredprod = item.benefittype;
                    newADBI.totalSumAssured += item.sumassured;
                    newADBI.totalNAR += item.reinsurednetamountatrisk;
                    lstADBI.Add(newADBI);
                    tableBasicRider.ADBI = lstADBI;

                }
                else if(item.benefittype == "ADB-IND")
                {
                    newADBIND.rider = item.baserider;
                    newADBIND.insuredprod = item.benefittype;
                    newADBIND.totalSumAssured += item.sumassured;
                    newADBIND.totalNAR += item.reinsurednetamountatrisk;
                    lstADBIND.Add(newADBIND);
                    tableBasicRider.ADBIND = lstADBIND;

                }
                else if(item.benefittype == "ADBR-GRP")
                {
                    newADBRGRP.rider = item.baserider;
                    newADBRGRP.insuredprod = item.benefittype;
                    newADBRGRP.totalSumAssured += item.sumassured;
                    newADBRGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstADBRGRP.Add(newADBRGRP);
                    tableBasicRider.ADBRGRP = lstADBRGRP;

                }
                else if(item.benefittype == "ADD&D-IND")
                {
                    newADDDIND.rider = item.baserider;
                    newADDDIND.insuredprod = item.benefittype;
                    newADDDIND.totalSumAssured += item.sumassured;
                    newADDDIND.totalNAR += item.reinsurednetamountatrisk;
                    lstADDDIND.Add(newADDDIND);
                    tableBasicRider.ADDDIND = lstADDDIND;

                }
                else if(item.benefittype == "ADP-GRP")
                {
                    newADPGRP.rider = item.baserider;
                    newADPGRP.insuredprod = item.benefittype;
                    newADPGRP.totalSumAssured += item.sumassured;
                    newADPGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstADPGRP.Add(newADPGRP);
                    tableBasicRider.ADPGRP = lstADPGRP;

                }
                else if(item.benefittype == "ADP-IND")
                {
                    newADPIND.rider = item.baserider;
                    newADPIND.insuredprod = item.benefittype;
                    newADPIND.totalSumAssured += item.sumassured;
                    newADPIND.totalNAR += item.reinsurednetamountatrisk;
                    lstADPIND.Add(newADPIND);
                    tableBasicRider.ADPIND = lstADPIND;

                }
                else if(item.benefittype == "BB-GRP")
                {
                    newBBGRP.rider = item.baserider;
                    newBBGRP.insuredprod = item.benefittype;
                    newBBGRP.totalSumAssured += item.sumassured;
                    newBBGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstBBGRP.Add(newBBGRP);
                    tableBasicRider.BBGRP = lstBBGRP;

                }
                else if(item.benefittype == "CIENDSRIND")
                {
                    newCIENDSRIND.rider = item.baserider;
                    newCIENDSRIND.insuredprod = item.benefittype;
                    newCIENDSRIND.totalSumAssured += item.sumassured;
                    newCIENDSRIND.totalNAR += item.reinsurednetamountatrisk;
                    lstCIENDSRIND.Add(newCIENDSRIND);
                    tableBasicRider.CIENDSRIND = lstCIENDSRIND;

                }
                else if(item.benefittype == "CIESIND")
                {
                    newCIESIND.rider = item.baserider;
                    newCIESIND.insuredprod = item.benefittype;
                    newCIESIND.totalSumAssured += item.sumassured;
                    newCIESIND.totalNAR += item.reinsurednetamountatrisk;
                    lstCCIESIND.Add(newCIESIND);
                    tableBasicRider.CIESIND = lstCCIESIND;

                }
                else if(item.benefittype == "CIRACIND")
                {
                    newCIRACIND.rider = item.baserider;
                    newCIRACIND.insuredprod = item.benefittype;
                    newCIRACIND.totalSumAssured += item.sumassured;
                    newCIRACIND.totalNAR += item.reinsurednetamountatrisk;
                    lstCIRACIND.Add(newCIRACIND);
                    tableBasicRider.CIRACIND = lstCIRACIND;

                }
                else if(item.benefittype == "CIRNAIND")
                {
                    newCIRNAIND.rider = item.baserider;
                    newCIRNAIND.insuredprod = item.benefittype;
                    newCIRNAIND.totalSumAssured += item.sumassured;
                    newCIRNAIND.totalNAR += item.reinsurednetamountatrisk;
                    lstCIRNAIND.Add(newCIRNAIND);
                    tableBasicRider.CIRNAIND = lstCIRACIND;

                }
                else if(item.benefittype == "CRITICALILLNESS")
                {
                    newCRITICALILLNESS.rider = item.baserider;
                    newCRITICALILLNESS.insuredprod = item.benefittype;
                    newCRITICALILLNESS.totalSumAssured += item.sumassured;
                    newCRITICALILLNESS.totalNAR += item.reinsurednetamountatrisk;
                    lstCRITICALILLNESS.Add(newCRITICALILLNESS);
                    tableBasicRider.CRITICALILLNESS = lstCRITICALILLNESS;
                }
                else if(item.benefittype == "DHIACCIND")
                {
                    newDHIACCIND.rider = item.baserider;
                    newDHIACCIND.insuredprod = item.benefittype;
                    newDHIACCIND.totalSumAssured += item.sumassured;
                    newDHIACCIND.totalNAR += item.reinsurednetamountatrisk;
                    lstDHIACCIND.Add(newDHIACCIND);
                    tableBasicRider.DHIACCIND = lstDHIACCIND;
                }
                else if(item.benefittype == "DHIBACGRP")
                {
                    newDHIBACGRP.rider = item.baserider;
                    newDHIBACGRP.insuredprod = item.benefittype;
                    newDHIBACGRP.totalSumAssured += item.sumassured;
                    newDHIBACGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstDHIBACGRP.Add(newDHIBACGRP);
                    tableBasicRider.DHIBACGRP = lstDHIBACGRP;
                }
                else if(item.benefittype == "DHIBALLGRP")
                {
                    newDHIBALLGRP.rider = item.baserider;
                    newDHIBALLGRP.insuredprod = item.benefittype;
                    newDHIBALLGRP.totalSumAssured += item.sumassured;
                    newDHIBALLGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstDHIBALLGRP.Add(newDHIBALLGRP);
                    tableBasicRider.DHIBALLGRP = lstDHIBALLGRP;
                }
                else if(item.benefittype == "DHIBALLIND")
                {
                    newDHIBALLIND.rider = item.baserider;
                    newDHIBALLIND.insuredprod = item.benefittype;
                    newDHIBALLIND.totalSumAssured += item.sumassured;
                    newDHIBALLIND.totalNAR += item.reinsurednetamountatrisk;
                    lstDHIBALLIND.Add(newDHIBALLIND);
                    tableBasicRider.DHIBALLIND = lstDHIBALLIND;
                }
                else if(item.benefittype == "DHIBILIND")
                {
                    newDHIBILIND.rider = item.baserider;
                    newDHIBILIND.insuredprod = item.benefittype;
                    newDHIBILIND.totalSumAssured += item.sumassured;
                    newDHIBILIND.totalNAR += item.reinsurednetamountatrisk;
                    lstDHIBILIND.Add(newDHIBILIND);
                    tableBasicRider.DHIBILIND = lstDHIBILIND;
                }
                else if(item.benefittype == "DOUBLEINDIND")
                {
                    newDOUBLEINDIND.rider = item.baserider;
                    newDOUBLEINDIND.insuredprod = item.benefittype;
                    newDOUBLEINDIND.totalSumAssured += item.sumassured;
                    newDOUBLEINDIND.totalNAR += item.reinsurednetamountatrisk;
                    lstDOUBLEINDIND.Add(newDOUBLEINDIND);
                    tableBasicRider.DOUBLEINDIND = lstDOUBLEINDIND;
                }
                else if(item.benefittype == "MCFRAIND")
                {
                    newMCFRAIND.rider = item.baserider;
                    newMCFRAIND.insuredprod = item.benefittype;
                    newMCFRAIND.totalSumAssured += item.sumassured;
                    newMCFRAIND.totalNAR += item.reinsurednetamountatrisk;
                    lstMCFRAIND.Add(newMCFRAIND);
                    tableBasicRider.MCFRAIND = lstMCFRAIND;
                }
                else if(item.benefittype == "MCFRININD")
                {
                    newMCFRININD.rider = item.baserider;
                    newMCFRININD.insuredprod = item.benefittype;
                    newMCFRININD.totalSumAssured += item.sumassured;
                    newMCFRININD.totalNAR += item.reinsurednetamountatrisk;
                    lstMCFRININD.Add(newMCFRININD);
                    tableBasicRider.MCFRININD = lstMCFRININD;
                }
                else if(item.benefittype == "MEDICALREIIND")
                {
                    newMEDICALREIIND.rider = item.baserider;
                    newMEDICALREIIND.insuredprod = item.benefittype;
                    newMEDICALREIIND.totalSumAssured += item.sumassured;
                    newMEDICALREIIND.totalNAR += item.reinsurednetamountatrisk;
                    lstMEDICALREIIND.Add(newMEDICALREIIND);
                    tableBasicRider.MEDICALREIIND = lstMEDICALREIIND;
                }
                else if(item.benefittype == "MEDICALREIMBURS")
                {
                    newMEDICALREIMBURS.rider = item.baserider;
                    newMEDICALREIMBURS.insuredprod = item.benefittype;
                    newMEDICALREIMBURS.totalSumAssured += item.sumassured;
                    newMEDICALREIMBURS.totalNAR += item.reinsurednetamountatrisk;
                    lstMEDICALREIMBURS.Add(newMEDICALREIMBURS);
                    tableBasicRider.MEDICALREIMBURS = lstMEDICALREIMBURS;
                }
                else if(item.benefittype == "MEDICALREIMGRP")
                {
                    newMEDICALREIMGRP.rider = item.baserider;
                    newMEDICALREIMGRP.insuredprod = item.benefittype;
                    newMEDICALREIMGRP.totalSumAssured += item.sumassured;
                    newMEDICALREIMGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstMEDICALREIMGRP.Add(newMEDICALREIMGRP);
                    tableBasicRider.MEDICALREIMGRP = lstMEDICALREIMGRP;
                }
                else if(item.benefittype == "MEDI-INS-IND")
                {
                    newMEDIINSIND.rider = item.baserider;
                    newMEDIINSIND.insuredprod = item.benefittype;
                    newMEDIINSIND.totalSumAssured += item.sumassured;
                    newMEDIINSIND.totalNAR += item.reinsurednetamountatrisk;
                    lstMEDIINSIND.Add(newMEDIINSIND);
                    tableBasicRider.MEDIINSIND = lstMEDIINSIND;
                }
                else if(item.benefittype == "MORTGAGEREDEMPT")
                {
                    newMORTGAGEREDEMPT.rider = item.baserider;
                    newMORTGAGEREDEMPT.insuredprod = item.benefittype;
                    newMORTGAGEREDEMPT.totalSumAssured += item.sumassured;
                    newMORTGAGEREDEMPT.totalNAR += item.reinsurednetamountatrisk;
                    lstMORTGAGEREDEMPT.Add(newMORTGAGEREDEMPT);
                    tableBasicRider.MORTGAGEREDEMPT = lstMORTGAGEREDEMPT;
                }
                else if(item.benefittype == "MRBACCIND")
                {
                    newMRBACCIND.rider = item.baserider;
                    newMRBACCIND.insuredprod = item.benefittype;
                    newMRBACCIND.totalSumAssured += item.sumassured;
                    newMRBACCIND.totalNAR += item.reinsurednetamountatrisk;
                    lstMRBACCIND.Add(newMRBACCIND);
                    tableBasicRider.MRBACCIND = lstMRBACCIND;
                }
                else if(item.benefittype == "MRPGRP")
                {
                    newMRPGRP.rider = item.baserider;
                    newMRPGRP.insuredprod = item.benefittype;
                    newMRPGRP.totalSumAssured += item.sumassured;
                    newMRPGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstMRPGRP.Add(newMRPGRP);
                    tableBasicRider.MRPGRP = lstMRPGRP;
                }
                else if(item.benefittype == "MURDER&ASSAULT-")
                {
                    newMURDERASSAULT.rider = item.baserider;
                    newMURDERASSAULT.insuredprod = item.benefittype;
                    newMURDERASSAULT.totalSumAssured += item.sumassured;
                    newMURDERASSAULT.totalNAR += item.reinsurednetamountatrisk;
                    lstMURDERASSAULT.Add(newMURDERASSAULT);
                    tableBasicRider.MURDERASSAULT = lstMURDERASSAULT;
                }
                else if(item.benefittype == "P&TDIS-INCO-IND")
                {
                    newPTDISINCOIND.rider = item.baserider;
                    newPTDISINCOIND.insuredprod = item.benefittype;
                    newPTDISINCOIND.totalSumAssured += item.sumassured;
                    newPTDISINCOIND.totalNAR += item.reinsurednetamountatrisk;
                    lstPTDISINCOIND.Add(newPTDISINCOIND);
                    tableBasicRider.PTDISINCOIND = lstPTDISINCOIND;
                }
                else if(item.benefittype == "RENEWALPERSONAL")
                {
                    newRENEWALPERSONAL.rider = item.baserider;
                    newRENEWALPERSONAL.insuredprod = item.benefittype;
                    newRENEWALPERSONAL.totalSumAssured += item.sumassured;
                    newRENEWALPERSONAL.totalNAR += item.reinsurednetamountatrisk;
                    lstRENEWALPERSONAL.Add(newRENEWALPERSONAL);
                    tableBasicRider.RENEWALPERSONAL = lstRENEWALPERSONAL;
                }
                else if(item.benefittype == "RPAR")
                {
                    newRPAR.rider = item.baserider;
                    newRPAR.insuredprod = item.benefittype;
                    newRPAR.totalSumAssured += item.sumassured;
                    newRPAR.totalNAR += item.reinsurednetamountatrisk;
                    lstRPAR.Add(newRPAR);
                    tableBasicRider.RPAR = lstRPAR;
                    }
                else if(item.benefittype == "SACIENDSTAPIND")
                    {
                        newSACIENDSTAPIND.rider = item.baserider;
                        newSACIENDSTAPIND.insuredprod = item.benefittype;
                        newSACIENDSTAPIND.totalSumAssured += item.sumassured;
                        newSACIENDSTAPIND.totalNAR += item.reinsurednetamountatrisk;
                        lstSACIENDSTAPIND.Add(newSACIENDSTAPIND);
                        tableBasicRider.SACIENDSTAPIND = lstSACIENDSTAPIND;
                    }
                else if(item.benefittype == "SACIESPIND")
                {
                    newSACIESPIND.rider = item.baserider;
                    newSACIESPIND.insuredprod = item.benefittype;
                    newSACIESPIND.totalSumAssured += item.sumassured;
                    newSACIESPIND.totalNAR += item.reinsurednetamountatrisk;
                    lstSACIESPIND.Add(newSACIESPIND);
                    tableBasicRider.SACIESPIND = lstSACIESPIND;
                }
                else if(item.benefittype == "SADB-IND")
                {
                    newSADBIND.rider = item.baserider;
                    newSADBIND.insuredprod = item.benefittype;
                    newSADBIND.totalSumAssured += item.sumassured;
                    newSADBIND.totalNAR += item.reinsurednetamountatrisk;
                    lstSADBIND.Add(newSADBIND);
                    tableBasicRider.SADBIND = lstSADBIND;
                }
                else if(item.benefittype == "SPLADBIND")
                {
                    newSPLADBIND.rider = item.baserider;
                    newSPLADBIND.insuredprod = item.benefittype;
                    newSPLADBIND.totalSumAssured += item.sumassured;
                    newSPLADBIND.totalNAR += item.reinsurednetamountatrisk;
                    lstSPLADBIND.Add(newSPLADBIND);
                    tableBasicRider.SPLADBIND = lstSPLADBIND;
                }
                else if(item.benefittype == "STANDALONECRITI")
                {
                    newSTANDALONECRITI.rider = item.baserider;
                    newSTANDALONECRITI.insuredprod = item.benefittype;
                    newSTANDALONECRITI.totalSumAssured += item.sumassured;
                    newSTANDALONECRITI.totalNAR += item.reinsurednetamountatrisk;
                    lstSTANDALONECRITI.Add(newSTANDALONECRITI);
                    tableBasicRider.STANDALONECRITI = lstSTANDALONECRITI;
                }
                else if(item.benefittype == "STANDALONEENH")
                {
                    newSTANDALONEENH.rider = item.baserider;
                    newSTANDALONEENH.insuredprod = item.benefittype;
                    newSTANDALONEENH.totalSumAssured += item.sumassured;
                    newSTANDALONEENH.totalNAR += item.reinsurednetamountatrisk;
                    lstSTANDALONEENH.Add(newSTANDALONEENH);
                    tableBasicRider.STANDALONEENH = lstSTANDALONEENH;
                }
                else if(item.benefittype == "T&PD-INCOME-GRP")
                {
                    newTPDINCOMEGRP.rider = item.baserider;
                    newTPDINCOMEGRP.insuredprod = item.benefittype;
                    newTPDINCOMEGRP.totalSumAssured += item.sumassured;
                    newTPDINCOMEGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstTPDINCOMEGRP.Add(newTPDINCOMEGRP);
                    tableBasicRider.TPDINCOMEGRP = lstTPDINCOMEGRP;
                }
                else if(item.benefittype == "T&PD-LS-GRP")
                {
                    newTPDLSGRP.rider = item.baserider;
                    newTPDLSGRP.insuredprod = item.benefittype;
                    newTPDLSGRP.totalSumAssured += item.sumassured;
                    newTPDLSGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstTPDLSGRP.Add(newTPDLSGRP);
                    tableBasicRider.TPDLSGRP = lstTPDLSGRP;
                }
                else if(item.benefittype == "T&PD-LS-IND")
                {
                    newTPDLSIND.rider = item.baserider;
                    newTPDLSIND.insuredprod = item.benefittype;
                    newTPDLSIND.totalSumAssured += item.sumassured;
                    newTPDLSIND.totalNAR += item.reinsurednetamountatrisk;
                    lstTPDLSIND.Add(newTPDLSIND);
                    tableBasicRider.TPDLSIND = lstTPDLSIND;
                }
                else if(item.benefittype == "T&TDIS-INCO-IND")
                {
                    newTTDISINCOIND.rider = item.baserider;
                    newTTDISINCOIND.insuredprod = item.benefittype;
                    newTTDISINCOIND.totalSumAssured += item.sumassured;
                    newTTDISINCOIND.totalNAR += item.reinsurednetamountatrisk;
                    lstTTDISINCOIND.Add(newTTDISINCOIND);
                    tableBasicRider.TTDISINCOIND = lstTTDISINCOIND;
                }
                else if(item.benefittype == "T&TDIS-LS-GRP")
                {
                    newTTDISLSGRP.rider = item.baserider;
                    newTTDISLSGRP.insuredprod = item.benefittype;
                    newTTDISLSGRP.totalSumAssured += item.sumassured;
                    newTTDISLSGRP.totalNAR += item.reinsurednetamountatrisk;
                    lstTTDISLSGRP.Add(newTTDISLSGRP);
                    tableBasicRider.TTDISLSGRP = lstTTDISLSGRP;
                }
                else if(item.benefittype == "T&TD-LS-IND")
                {
                    newTTDLSIND.rider = item.baserider;
                    newTTDLSIND.insuredprod = item.benefittype;
                    newTTDLSIND.totalSumAssured += item.sumassured;
                    newTTDLSIND.totalNAR += item.reinsurednetamountatrisk;
                    lstTTDLSIND.Add(newTTDLSIND);
                    tableBasicRider.TTDLSIND = lstTTDLSIND;
                }
                else if(item.benefittype == "TEMRIDISNIND")
                {
                    newTEMRIDISNIND.rider = item.baserider;
                    newTEMRIDISNIND.insuredprod = item.benefittype;
                    newTEMRIDISNIND.totalSumAssured += item.sumassured;
                    newTEMRIDISNIND.totalNAR += item.reinsurednetamountatrisk;
                    lstTEMRIDISNIND.Add(newTEMRIDISNIND);
                    tableBasicRider.TEMRIDISNIND = lstTEMRIDISNIND;
                }
                else if(item.benefittype == "TERMRIDER(PAYOR")
                {
                    newTERMRIDERPAYOR.rider = item.baserider;
                    newTERMRIDERPAYOR.insuredprod = item.benefittype;
                    newTERMRIDERPAYOR.totalSumAssured += item.sumassured;
                    newTERMRIDERPAYOR.totalNAR += item.reinsurednetamountatrisk;
                    lstTERMRIDERPAYOR.Add(newTERMRIDERPAYOR);
                    tableBasicRider.TERMRIDERPAYOR = lstTERMRIDERPAYOR;
                }
                else if(item.benefittype == "TIR")
                {
                    newTIR.rider = item.baserider;
                    newTIR.insuredprod = item.benefittype;
                    newTIR.totalSumAssured += item.sumassured;
                    newTIR.totalNAR += item.reinsurednetamountatrisk;
                    lstTIR.Add(newTIR);
                    tableBasicRider.TIR = lstTIR;
                }
                else if(item.benefittype == "VARLIFE-GU")
                {
                    newVARLIFEGU.rider = item.baserider;
                    newVARLIFEGU.insuredprod = item.benefittype;
                    newVARLIFEGU.totalSumAssured += item.sumassured;
                    newVARLIFEGU.totalNAR += item.reinsurednetamountatrisk;
                    lstVARLIFEGU.Add(newVARLIFEGU);
                    tableBasicRider.VARLIFEGU = lstVARLIFEGU;
                }
                else if(item.benefittype == "WOP-D&DP-IND")
                {
                    newWOPDDPIND.rider = item.baserider;
                    newWOPDDPIND.insuredprod = item.benefittype;
                    newWOPDDPIND.totalSumAssured += item.sumassured;
                    newWOPDDPIND.totalNAR += item.reinsurednetamountatrisk;
                    lstWOPDDPIND.Add(newWOPDDPIND);
                    tableBasicRider.WOPDDPIND = lstWOPDDPIND;
                }
                else if(item.benefittype == "WOP-DDI-IND")
                {
                    newWOPDDIIND.rider = item.baserider;
                    newWOPDDIIND.insuredprod = item.benefittype;
                    newWOPDDIIND.totalSumAssured += item.sumassured;
                    newWOPDDIIND.totalNAR += item.reinsurednetamountatrisk;
                    lstWOPDDIIND.Add(newWOPDDIIND);
                    tableBasicRider.WOPDDIIND = lstWOPDDIIND;
                }
                else if(item.benefittype == "WOPDIIND")
                {
                    newWOPDIIND.rider = item.baserider;
                    newWOPDIIND.insuredprod = item.benefittype;
                    newWOPDIIND.totalSumAssured += item.sumassured;
                    newWOPDIIND.totalNAR += item.reinsurednetamountatrisk;
                    lstWOPDIIND.Add(newWOPDIIND);
                    tableBasicRider.WOPDIIND = lstWOPDIIND;
                }
                else if(item.benefittype == "WOP-DI-IND")
                {
                    newWOPDIIND_.rider = item.baserider;
                    newWOPDIIND_.insuredprod = item.benefittype;
                    newWOPDIIND_.totalSumAssured += item.sumassured;
                    newWOPDIIND_.totalNAR += item.reinsurednetamountatrisk;
                    lstWOPDIIND_.Add(newWOPDIIND_);
                    tableBasicRider.WOPDIIND_ = lstWOPDIIND_;
                }
                else if(item.benefittype == "WOPDOPIND")
                {
                    newWOPDOPIND.rider = item.baserider;
                    newWOPDOPIND.insuredprod = item.benefittype;
                    newWOPDOPIND.totalSumAssured += item.sumassured;
                    newWOPDOPIND.totalNAR += item.reinsurednetamountatrisk;
                    lstWOPDOPIND.Add(newWOPDOPIND);
                    tableBasicRider.WOPDOPIND = lstWOPDOPIND;
                }
                else if(item.benefittype == "WOP-DP-IND")
                {
                    newWOPDPIND.rider = item.baserider;
                    newWOPDPIND.insuredprod = item.benefittype;
                    newWOPDPIND.totalSumAssured += item.sumassured;
                    newWOPDPIND.totalNAR += item.reinsurednetamountatrisk;
                    lstWOPDPIND.Add(newWOPDPIND);
                    tableBasicRider.WOPDPIND = lstWOPDPIND;
                }
            }
            #endregion
            ViewBag.Identifier = Identifier;
            ViewBag.FullName = strFullname.ToUpper();
            ViewBag.TotalPolicy = intPolicyNo;
            ViewBag.DateofBirth = strDOB;

            return View("ViewAccumulation", tableBasicRider);
        }


            
        public IActionResult ImportToExcel(string Identifier)
        {
            var queryBasics = _db.dbLifeData.Where(y => y.identifier == Identifier && y.baserider == "BASIC");
            var queryRiders = _db.dbLifeData.Where(y => y.identifier == Identifier && y.baserider == "RIDER");
            
            var lstAccXML= new List<AccumulationXML>();
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Accumulation");
            System.IO.Stream spreadsheetStream = new System.IO.MemoryStream();
            Stream fs = new MemoryStream();

            DataTable dt = new DataTable();
            decimal dclTotalSAR = 0;
            decimal dclTotalNAR = 0;
            string strFullname = string.Empty;
            string strDOB = string.Empty;
            int rowcount = 1;
            
            foreach(var item in queryBasics) //Accumulate Basics
            {
                var newAccumulationXML = new AccumulationXML();
                newAccumulationXML.POLICY_NO = item.policyno;
                newAccumulationXML.BENEFIT_TYPE = item.benefittype;
                newAccumulationXML.QUARTER = item.soaperiod;
                newAccumulationXML.BASIC_RIDER = item.baserider;
                strFullname = item.fullName;
                strDOB = item.dateofbirth;
                newAccumulationXML.BORDEREAUX_YEAR = Convert.ToInt32(item.bordereauxyear);
                newAccumulationXML.BORDEREAUX_FILENAME = item.bordereauxfilename;
                newAccumulationXML.TOTAL_SUM_ASSURED = item.sumassured;
                newAccumulationXML.TOTAL_NET_AMOUNT_AT_RISK = item.reinsurednetamountatrisk;
                dclTotalSAR += item.sumassured;
                dclTotalNAR += item.reinsurednetamountatrisk;
                rowcount ++;
                lstAccXML.Add(newAccumulationXML);
            }
            ws.Cell(1, 1).Value = ("FullName :");
            ws.Cell(1, 2).Value = strFullname;
            ws.Cell(2, 1).Value = ("Date of Birth :");
            ws.Cell(2, 2).Value = strDOB;
            ws.Cell(4, 1).Value = ("ACCUMULATION");
            ws.Cell(5, 1).InsertTable(lstAccXML);
            ws.Cell(rowcount + 5, 7).Value = dclTotalSAR;
            ws.Cell(rowcount + 5, 8).Value = dclTotalNAR;

            fs.Position = 0;

            using(MemoryStream stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Accumulation Report(" + Identifier + ")"+ DateTime.Now.ToString("MM-dd-yyyy") + ".xlsx");
            }


        }
    }

}
