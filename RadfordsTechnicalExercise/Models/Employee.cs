using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity;
using System.Web.Mvc;
using System.Linq;
using System.Web;
using System.Reflection;

namespace RadfordsTechnicalExercise.Models
{
	public class Employee
	{
		public int Id { get; set; }

		[DisplayName( "First Name" )]
		[Required]
		public string FirstName { get; set; }

		[DisplayName( "Middle Initial" )]
		[StringLength( 1 , ErrorMessage = "Middle initial exceeds the maximum length of {1} character" )]
		public string MiddleInitial { get; set; }

		[DisplayName( "Last Name" )]
		[Required]
		public string LastName { get; set; }

		[DisplayName( "Home Phone" )]
		[StringLength( 30 , ErrorMessage = "Home phone number exceeds the maximum length of {1} characters" )]
		public string HomePhoneNumber { get; set; }

		[DisplayName( "Cell Phone" )]
		[StringLength( 30 , ErrorMessage = "Cell phone number exceeds the maximum length of {1} characters" )]
		public string CellPhoneNumber { get; set; }

		[DisplayName( "Office Extension" )]
		[StringLength( 10 , ErrorMessage = "Office extension number exceeds the maximum length of {1} characters" )]
		public string OfficeExtension { get; set; }

		[DisplayName( "IRD Number" )]
		[StringLength( 9 , ErrorMessage = "IRD numbers exceeds the maximum length of {1} characters" )]
		public string TaxNumber { get; set; }

		[UIHint( "Checkbox" )]
		[DisplayName( "Employee is active" )]
		public bool IsActive { get; set; }

		public void ValidateEmployee( ModelStateDictionary modelState , DbSet<Employee> dbSet, Employee employee )
		{
			if ( modelState.IsValid )
			{
				var listOfInputsToCheck = new Dictionary<string, string> {
					{ "HomePhoneNumber", employee.HomePhoneNumber },
					{ "CellPhoneNumber", employee.CellPhoneNumber },
					{ "OfficeExtension", employee.OfficeExtension },
					{ "TaxNumber", employee.TaxNumber }
				};

				// validate number fields
				foreach ( var input in listOfInputsToCheck )
				{
					if ( !string.IsNullOrEmpty( input.Value ) )
					{
						if ( !IsValidNumber( input.Value ) )
						{
							modelState.AddModelError( input.Key , "This field can only contain numbers" );
						}
					}
				}
				
				// only allow max of 15 entries in DB at
				if ( dbSet.Count( ) == 15 )
				{
					modelState.AddModelError( "" , "You can only store a maximum of 15 employees' details at this time" );
				}
			}
		}

		public bool IsValidNumber( string inputToCheck )
		{
			int parsedNumber;

			return int.TryParse( inputToCheck , out parsedNumber );
		}

	}



	public class EmployeeDBContext : DbContext
	{
		public DbSet<Employee> Employees { get; set; }
	}
}