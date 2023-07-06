import openpyxl

def validation(arrFields):
    validation_code = []
    validation_code.append('''public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {''')
    for x in arrFields[1:]:
        if x[1] == 'string':
            if x[6] == None and x[7] == None:
                if x[8] is not None and x[9] == 'phone' :
                    validation_code.append('''//'''+x[0]+'''
            if ('''+x[0]+''' != null && ('''+x[0]+'''_validation.Length == '''+str(x[8])+'''))
            {
                yield return new ValidationResult('Invalid '''+x[0]+''' ');    
            }''')
                elif x[8] == None and x[9] == 'email':
                    validation_code.append('''//'''+x[0]+'''
            if ('''+x[0]+''' != null && !Regex.IsMatch('''+x[0]+'''_validation.ToLower(), @"^.{0,14}[^@.]*$"))
            {
                yield return new ValidationResult('Invalid '''+x[0]+''' ');
            }''')
            elif x[7] == None:
                validation_code.append('''//'''+x[0]+'''
            if ('''+x[0]+''' != null && ('''+x[0]+'''_validation.Length < '''+str(x[6])+'''))
            {
                yield return new ValidationResult('Invalid '''+x[0]+''' ');    
            }''')
            elif x[6] == None:
                validation_code.append('''//'''+x[0]+'''
            if ('''+x[0]+''' != null && ('''+x[0]+'''_validation.Length > '''+str(x[7])+'''))
            {
                yield return new ValidationResult('Invalid '''+x[0]+''' ');    
            }''')
            elif x[9] == 'textarea': 
                validation_code.append('''//'''+x[0]+'''
            if ('''+x[0]+''' != null && ('''+x[0]+'''_validation.Length < '''+str(x[6])+''' || '''+x[0]+'''_validation.Length > '''+str(x[7])+'''))
            {
                yield return new ValidationResult('Invalid '''+x[0]+''' ');    
            }''')
            elif x[9] == 'text':
                validation_code.append('''//'''+x[0]+'''
            if ('''+x[0]+''' != null && ('''+x[0]+'''_validation.Length < '''+str(x[6])+''' || '''+x[0]+'''_validation.Length > '''+str(x[7])+'''))
            {
                yield return new ValidationResult('Invalid '''+x[0]+''' ');    
            }''')

        elif x[1] == 'int':
            if x[4] == None and x[5] == None:
                validation_code.append('''//'''+x[0]+'''
            No value found''')
            elif x[4] == None:
                validation_code.append('''//'''+x[0]+'''
            if ('''+x[0]+''' > '''+str(x[5])+''')
            {
                yield return new ValidationResult('Invalid '''+x[0]+''' ');
            }''')
            elif x[5] == None:
                validation_code.append('''//'''+x[0]+'''
            if ('''+x[0]+''' < '''+str(x[4])+''')
            {
                yield return new ValidationResult('Invalid '''+x[0]+''' ');
            }''')
            else:
                validation_code.append('''//'''+x[0]+'''
            if ('''+x[0]+''' <= '''+str(x[4])+''' || '''+x[0]+''' >= '''+str(x[5])+''')
            {
                yield return new ValidationResult('Invalid '''+x[0]+''' ');
            }''')
                
        elif x[1] == 'date':
            validation_code.append('''//'''+x[0]+'''
            if (!IsValidDate('''+x[0]+'''))
            {
                yield return new ValidationResult('Invalid '''+x[0]+''' ');
            }
            //'''+x[0]+''' validation function
            static bool IsValidDate(string input)
            {
                return DateTime.TryParse(input, out _);
            }
        ''')

    validation_code = "\n\t\t\t".join(validation_code)
    return validation_code

def makeController(tableName, arrFields):
    def starter_template():
        start_code = '''using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using AnimalKatta.Api.Models;
using Microsoft.AspNetCore.Authorization;
using AnimalKatta.Api.ViewModel;
using System.Text.Json;
using AnimalKatta.Api.Services;
//using Newtonsoft.Json.Linq;

namespace AnimalKatta.Api.Controllers'''
        return start_code

    def initial_function():
        initial_function_code = '''\n\t\t
{
    //[Authorize] 
    [ApiController]
    public class '''+tableName+'''Controller : ControllerBase
    {
        private readonly nopcommerceContext _context;

        public '''+tableName+'''Controller(nopcommerceContext context)
        {
            _context = context;
        }'''
        
        return initial_function_code
    
    def List():
        List_code = '''\n\t\t
        //List all the records
        [Authorize]
        [Route("~/api/[controller]/'''+tableName+'''list")]
        [HttpPost]
        public IActionResult '''+tableName+'''List('''+tableName+''' req)
        {
            try
            {
                var skip = 0;  
                
                if(req.per_page)
                {
                    skip = (req.page_no - 1) * req.per_page;     
                    var result = (from cw in _context.'''+tableName+'''
                                select cw
                                ).Skip(skip).Take(req.per_page).toList();
                    return Ok(result);
                }    
                else
                {
                    var result = (from cw in _context.'''+tableName+'''
                                  select cw
                                  );
                    return Ok(result);
                }
            }
            catch (Exception e)
            {
                return StatusCode(300, "Exception message");
            }
        }'''
        return List_code
    
    def ListById():
        ListById_code = '''\n\t\t   
        //List the records by ID
        [Authorize]
        [Route("~/api/[controller]/'''+tableName+'''listById")]
        [HttpPost]
        public IActionResult '''+tableName+'''ListById('''+tableName+''' req)
        {
            try
            {
                var skip = 0;  
                
                if(req.per_page)
                {
                    skip = (req.page_no - 1) * req.per_page;     
                    var result = (from cw in _context.'''+tableName+'''
                                select cw
                                ).Skip(skip).Take(req.per_page).toList();
                    return Ok(result);
                }    
                else
                {
                    var result = (from cw in _context.'''+tableName+'''
                                  select cw
                                  );
                    return Ok(result);
                }
            
                var result = (from cw in _context.'''+tableName+'''
                                where %s
                                select cw
                                ).toList();
                return Ok(result);
            }
            catch (Exception e)
            {
                return StatusCode(300, "Exception message");
            }
        }'''
        
        And_array = []
        if len(Unique__fields) > 0:
            And_array.append("cw." + Unique__fields[0] + ' == ' + "req." + Unique__fields[0])

            if len(Unique__fields) > 1:
                for value in Unique__fields[1:]:
                    And_array.append('cw.' + value + ' == ' + 'req.' + value)

        output = ' && '.join(And_array) #join for concatenating values of array with &&
        ListById_code_1 = ListById_code % (output)
        return ListById_code_1
    
    def Add():
        Add_code = '''\n\t\t
        //Add the records
        [Authorize]
        [Route("~/api/[controller]/'''+tableName+'''Add")]
        [HttpPost]
        public IActionResult '''+tableName+'''Add('''+tableName+''' req)
        {
            try
            {
                var skip = 0;  
                
                if(req.per_page)
                {
                    skip = (req.page_no - 1) * req.per_page;     
                    var result = (from cw in _context.'''+tableName+'''
                                select cw
                                ).Skip(skip).Take(req.per_page).toList();
                    return Ok(result);
                }    
                else
                {
                    var result = (from cw in _context.'''+tableName+'''
                                  select cw
                                  );
                    return Ok(result);
                }

                '''+tableName+''' reqnew = new '''+tableName+'''();
            %s
                _context.'''+tableName+'''.Add(reqnew);
                _context.SaveChanges();
                return StatusCode(200, " '''+tableName+''' has been added successfully");
            }
            catch (Exception e)
            {
                return StatusCode(300, "Exception message");
            }
        }'''
        
        Add = []
        for x in arrFields:
            if x[3] == 'yes':
                Add.append('''reqnew.'''+x[0]+''' = req.'''+x[0]+''';''')
        Add = "\n\t\t\t".join("\t" + line for line in Add)
        Add_code_1 = Add_code % (Add)        
        
        
        return Add_code_1
        
    def Edit():
        Edit_code = '''\n\t\t
        //Edit the records
        [Authorize]
        [Route("~/api/[controller]/'''+tableName+'''Edit")]
        [HttpPost]
        public IActionResult '''+tableName+'''Edit('''+tableName+''' req)
        {
            try
            {
                var skip = 0;  
                
                if(req.per_page)
                {
                    skip = (req.page_no - 1) * req.per_page;     
                    var result = (from cw in _context.'''+tableName+'''
                                select cw
                                ).Skip(skip).Take(req.per_page).toList();
                    return Ok(result);
                }    
                else
                {
                    var result = (from cw in _context.'''+tableName+'''
                                  select cw
                                  );
                    return Ok(result);
                }
            
                var check_data = _context.'''+tableName+'''.FirstOrDefault(x => (%s))
                //return Ok(check_data);
                if (check_data != null)
                {
                %s
                    _context.SaveChanges();
                    return StatusCode(200, " '''+tableName+''' has been updated successfully");
                }
                else
                {
                    return StatusCode(300, "'''+tableName+''' details not found");
                }
            }
            catch (Exception e)
            {
                return StatusCode(300, "Exception message");
            }
        }'''
        
        
        And_array = []
        if len(Unique__fields) > 0:
            And_array.append("x." + Unique__fields[0] + ' == ' + "req." + Unique__fields[0])

            if len(Unique__fields) > 1:
                for value in Unique__fields[1:]:
                    And_array.append('x.' + value + ' == ' + 'req.' + value)

        output = ' && '.join(And_array) #join for concatenating values of array with &&

        Edit = []
        for x in arrFields:
            if x[3] == 'yes':
                Edit.append('''check_data.'''+x[0]+''' = req.'''+x[0]+''';''')
        Edit = "\n\t\t\t\t".join("\t" + line for line in Edit)

        Edit_code_1 = Edit_code % (output, Edit)
        
        return Edit_code_1
    
    def Delete():
        Delete_code = '''\n\t\t
        //Delete the records
        [Authorize]
        [Route("~/api/[controller]/Dele'''+tableName+''' ")]
        [HttpPost]
        public IActionResult Dele'''+tableName+'''('''+tableName+''' req)
        {
            try
            {
                var skip = 0;  
                
                if(req.per_page)
                {
                    skip = (req.page_no - 1) * req.per_page;     
                    var result = (from cw in _context.'''+tableName+'''
                                select cw
                                ).Skip(skip).Take(req.per_page).toList();
                    return Ok(result);
                }    
                else
                {
                    var result = (from cw in _context.'''+tableName+'''
                                  select cw
                                  );
                    return Ok(result);
                }

                var check_data = _context.'''+tableName+'''. FirstOrDefault(x => (%s))
                if (check_data != null)
                {
                    JsonOutptDataApp json = new JsonOutptDataApp();
                    json.status = true;
                    check_data.is_active=req.is_active;
                    _context.SaveChanges();
                    json.message = " '''+tableName+''' has been deleted successfully";
                    return Ok(json);
                }
                else
                {
                    return StatusCode(300, " '''+tableName+''' Details Not Found");
                }

            }
            catch (Exception e)
            {
                return StatusCode(300, "Exception message");
            }
        }
    }
}'''
        
        And_array = []
        if len(Unique__fields) > 0:
            And_array.append("x." + Unique__fields[0] + ' == ' + "req." + Unique__fields[0])

            if len(Unique__fields) > 1:
                for value in Unique__fields[1:]:
                    And_array.append('x.' + value + ' == ' + 'req.' + value)

        output = ' && '.join(And_array) #join for concatenating values of array with &&
        
        Delete_code_1 =Delete_code % (output)
        
        return Delete_code_1
        
    
    fun1= starter_template()
    fun2= initial_function()
    fun3= List()
    fun4= ListById()
    fun5= Add()
    fun6= Edit()
    fun7 = Delete()
    return (fun1+fun2+fun3+fun4+fun5+fun6+fun7)

def makeModel(tableName, arrFields):
    def starter_template_model():
        start_code_model = '''
using System;
using System.Collections.Generic;
using System.Linq;'''
        return start_code_model

        
    def model_function():
        model_function_code = '''
namespace ProjecName.Api.Models
{
    public partial class '''+tableName+'''
    {
        [key]
    %s  
        %s
        }
    }
}
        
'''
        
        model_code = []
        for x in arrFields[1:]:
            if x[1] == 'string':
                model_code.append('''public '''+x[1]+''' '''+x[0]+'''{ get; set; }''')
            elif x[1] == 'int' and x[2] == 'yes':
                model_code.append('''public '''+x[1]+''' '''+x[0]+'''{ get; set; }''')
            else: 
                model_code.append('''public '''+x[1]+'''? '''+x[0]+'''{ get; set; }''')
                
        model_code = "\n\t".join("\t" + line for line in model_code)
        
        validations = validation(arrFields)
        
        model_function_code_1 = model_function_code % (model_code, validations)
        
        
        return model_function_code_1
    
    
    mod1 = starter_template_model()
    mod2 = model_function()
    return (mod1+mod2)

def makeQueryModel(tableName, arrFields):
    def starter_template_QureyModel():
        start_Code_QueryModel = '''using System;
using System.Collections.Generic;


namespace ProjecName.Api.Models
{
    public partial class QueryModel
    {
        
        public int? page_no  {get;set;}
        public int? per_page  {get;set;}
    }
}
'''
        return start_Code_QueryModel
    query_mod1 = starter_template_QureyModel()
    
    return (query_mod1)
        




# Open the Excel file
workbook = openpyxl.load_workbook('dot NET.xlsx')

# Select the first sheet
sheet = workbook.active

# Variables for storing tableName and arrFields
tableName = None
arrFields = []


# Iterate over each row in the sheet
for row in sheet.iter_rows(values_only=True):
    cell1 = row[0]
    cell2 = row[1]
    if cell1 and not cell2:
        # Save the value of the first cell in tableName
        
        tableName = cell1

        
    elif cell1 and cell2:
        # Add the row as an array in arrFields
        arrFields.append(row)
        
        #Listing Unique Fields
        Unique__fields = []
        for x in arrFields:
            if x[2] == 'yes':
                Unique__fields.append(x[0])

        
        file_name = tableName+"_Controller.cs"  
        with open(file_name, 'w') as file:
            file.write(makeController(tableName, arrFields))
            
        file_name = "Warehouse.cs"  
        with open(file_name, 'w') as file:
            file.write(makeModel(tableName, arrFields))
            
        file_name = "QueryModel.cs"  
        with open(file_name, 'w') as file:
            file.write(makeQueryModel(tableName, arrFields))
        makeQueryModel(tableName, arrFields)

    elif not cell1:
        makeController(tableName, arrFields)
        makeModel(tableName, arrFields)
        makeQueryModel(tableName, arrFields)
        tableName = None
        arrFields = []

    
print("Current Table name is: ",tableName)

print("Current Table value is: ",arrFields)
