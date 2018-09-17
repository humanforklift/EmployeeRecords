using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using RadfordsTechnicalExercise.Models;
using ClosedXML.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;

namespace RadfordsTechnicalExercise.Controllers
{
    public class EmployeesController : Controller
    {
        private EmployeeDBContext db = new EmployeeDBContext();

        // GET: Employees
        public ActionResult Index()
        {
            return View(db.Employees.ToList());
        }

        // GET: Employees/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // GET: Employees/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Employees/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create( Employee employee)
        {
			employee.ValidateEmployee( ModelState , db.Employees, employee );

			if (ModelState.IsValid)
            {
                db.Employees.Add(employee);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(employee);
        }

        // GET: Employees/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // POST: Employees/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit( Employee employee)
        {
			employee.ValidateEmployee( ModelState , db.Employees , employee );

			if (ModelState.IsValid)
            {
                db.Entry(employee).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(employee);
        }

        // GET: Employees/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // POST: Employees/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Employee employee = db.Employees.Find(id);
            db.Employees.Remove(employee);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

		public void Export( )
		{
			var employee = new Employee( );
			var employeeFields = employee.GetType( ).GetProperties( );
			var database = db.Employees.ToArray();

			DataTable table = new DataTable( );

			foreach ( var field in employeeFields )
			{
				table.Columns.Add( field.Name , typeof( string ) );
			}

			foreach ( var item in database )
			{
				table.Rows.Add( 
					item.Id, 
					item.FirstName, 
					item.MiddleInitial, 
					item.LastName,
					item.HomePhoneNumber,
					item.CellPhoneNumber,
					item.OfficeExtension,
					item.TaxNumber,
					item.IsActive
					);
			}

			table.Columns.Remove( "Id" );


		var wbook = new XLWorkbook( );
			wbook.Worksheets.Add( table , "EmployeeData" );

			wbook.Worksheets.First( ).Columns( ).AdjustToContents( );

			HttpResponseBase httpResponse = Response;
			httpResponse.Clear( );
			httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			httpResponse.AddHeader( "content-disposition" , "attachment;filename=\"EmployeeContactList" + DateTime.Now.ToShortDateString() + ".xlsx\"" );

			using ( MemoryStream memoryStream = new MemoryStream( ) )
			{
				wbook.SaveAs( memoryStream );
				memoryStream.WriteTo( httpResponse.OutputStream );
				memoryStream.Close( );
			}

			httpResponse.End( );
		}

		public void Print( )
		{
			var document = new Document( PageSize.A4.Rotate( ) , 10 , 10 , 25 , 25 );

			var boldTableFont = FontFactory.GetFont( "Arial" , 10 , Font.BOLD );
			var bodyFont = FontFactory.GetFont( "Arial" , 10 , Font.NORMAL );

			var output = new MemoryStream( );
			var writer = PdfWriter.GetInstance( document , output );

			document.Open( );

			var employeeTable = new PdfPTable( 8 ) { WidthPercentage = 100 };
			employeeTable.HorizontalAlignment = 0;

			AddTableHeaders( employeeTable , boldTableFont );
			AddTableContent( employeeTable , bodyFont );

			document.Add( employeeTable );

			document.Close( );

			Response.ContentType = "application/pdf";
			Response.AddHeader( "Content-Disposition" , "attachment;filename=\"EmployeeContactList" + DateTime.Now.ToShortDateString( ) + ".pdf\"" );
			Response.BinaryWrite( output.ToArray( ) );

		}

		public float[ ] GetHeaderWidths( Font font , params string[ ] headers )
		{
			var cumulativeWidth = 0;
			var columns = headers.Length;
			var widthOfEachColumn = new int[ columns ];
			for ( var i = 0; i < columns; ++i )
			{
				var width = font.GetCalculatedBaseFont( true ).GetWidth( headers[ i ] );
				cumulativeWidth += width;
				widthOfEachColumn[ i ] = width;
			}
			var result = new float[ columns ];
			for ( var i = 0; i < columns; ++i )
			{
				result[ i ] = ( float ) widthOfEachColumn[ i ] / cumulativeWidth * 100;
			}
			return result;
		}

		public void AddTableHeaders(PdfPTable table, Font font)
		{
			var headers = new string[ ]
			{
				"First Name",
				"Middle Initial",
				"Last Name",
				"Home Phone",
				"Cell Phone",
				"Office Ext.",
				"IRD Number",
				"Is Active"
			};

			//table.DefaultCell.Border = 0;
			table.SetWidths( GetHeaderWidths( font , headers ) );

			for ( int i = 0; i < headers.Length; ++i )
			{
				table.AddCell( new Phrase( headers[ i ] , font ) );
			}
		} 

		public void AddTableContent( PdfPTable table , Font font )
		{
			var employees = db.Employees.Select( s => new {
				s.FirstName,
				s.MiddleInitial,
				s.LastName,
				s.HomePhoneNumber,
				s.CellPhoneNumber,
				s.OfficeExtension,
				s.TaxNumber,
				s.IsActive
			} ).ToList( );

			foreach ( var employee in employees )
			{
				table.AddCell( new Phrase( employee.FirstName , font ) );
				table.AddCell( new Phrase( employee.MiddleInitial , font ) );
				table.AddCell( new Phrase( employee.LastName , font ) );
				table.AddCell( new Phrase( employee.HomePhoneNumber , font ) );
				table.AddCell( new Phrase( employee.CellPhoneNumber , font ) );
				table.AddCell( new Phrase( employee.OfficeExtension , font ) );
				table.AddCell( new Phrase( employee.TaxNumber , font ) );
				table.AddCell( new Phrase( employee.IsActive ? "True" : "False" , font ) );
			}

		}

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
