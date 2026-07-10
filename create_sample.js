const XLSX = require('xlsx');

const workbook = XLSX.utils.book_new();
const data = [
  { "Employee ID": "EMP001", "First Name": "John", "Last Name": "Doe", "Department": "Engineering", "Position": "Software Engineer", "Email": "john.doe@company.com", "Phone": "555-0101", "Salary": 75000 },
  { "Employee ID": "EMP002", "First Name": "Jane", "Last Name": "Smith", "Department": "Marketing", "Position": "Marketing Manager", "Email": "jane.smith@company.com", "Phone": "555-0102", "Salary": 85000 },
  { "Employee ID": "EMP003", "First Name": "Bob", "Last Name": "Johnson", "Department": "HR", "Position": "HR Specialist", "Email": "bob.johnson@company.com", "Phone": "555-0103", "Salary": 65000 },
  { "Employee ID": "EMP004", "First Name": "Alice", "Last Name": "Williams", "Department": "Finance", "Position": "Accountant", "Email": "alice.williams@company.com", "Phone": "555-0104", "Salary": 70000 },
  { "Employee ID": "EMP005", "First Name": "Charlie", "Last Name": "Brown", "Department": "Engineering", "Position": "Senior Developer", "Email": "charlie.brown@company.com", "Phone": "555-0105", "Salary": 95000 },
  { "Employee ID": "EMP006", "First Name": "Diana", "Last Name": "Davis", "Department": "Sales", "Position": "Sales Representative", "Email": "diana.davis@company.com", "Phone": "555-0106", "Salary": 60000 },
  { "Employee ID": "EMP007", "First Name": "Edward", "Last Name": "Miller", "Department": "IT", "Position": "System Administrator", "Email": "edward.miller@company.com", "Phone": "555-0107", "Salary": 80000 },
  { "Employee ID": "EMP008", "First Name": "Fiona", "Last Name": "Wilson", "Department": "Design", "Position": "UX Designer", "Email": "fiona.wilson@company.com", "Phone": "555-0108", "Salary": 78000 },
  { "Employee ID": "EMP009", "First Name": "George", "Last Name": "Moore", "Department": "Engineering", "Position": "DevOps Engineer", "Email": "george.moore@company.com", "Phone": "555-0109", "Salary": 90000 },
  { "Employee ID": "EMP010", "First Name": "Hannah", "Last Name": "Taylor", "Department": "Operations", "Position": "Operations Manager", "Email": "hannah.taylor@company.com", "Phone": "555-0110", "Salary": 88000 },
  { "Employee ID": "EMP011", "First Name": "Ivan", "Last Name": "Anderson", "Department": "Legal", "Position": "Legal Counsel", "Email": "ivan.anderson@company.com", "Phone": "555-0111", "Salary": 110000 },
  { "Employee ID": "EMP012", "First Name": "Julia", "Last Name": "Thomas", "Department": "Support", "Position": "Support Lead", "Email": "julia.thomas@company.com", "Phone": "555-0112", "Salary": 72000 },
];

const worksheet = XLSX.utils.json_to_sheet(data);
XLSX.utils.book_append_sheet(workbook, worksheet, "Employees");
XLSX.writeFile(workbook, "sample_employees.xlsx");
console.log("Created sample_employees.xlsx");
