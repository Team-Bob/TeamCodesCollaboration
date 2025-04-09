
//Imports for handling excel, input, date and time formatting
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Scanner;
import java.util.*;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.LocalDate;
import java.time.format.DateTimeParseException;
import java.text.DecimalFormat;


public class PayrollSystemForMotorPH {
    
// Combining All the codes from Task 7 to Task 10
// MS2 Initial Code is Hardcoded 
// MS 2 Final Code    
    
    public static void main(String[] args) {
        Scanner s = new Scanner(System.in);

        
// For User input (Employee Number)
    System.out.print("Enter Employee Number: ");
        int employeeNumberInput = s.nextInt(); 

// File path to the Excel database
    String filePath = "src//MotorPH Employee Data .xlsx";
        Map<String, String[]> employeeData = getEmployeeDetails(filePath,employeeNumberInput);
       
// If the employee is not found it will not ask a date for attendance   
        if (employeeData.isEmpty()) {
        System.out.println("Employee Number not found.");
        return; // Stop execution if employee not found
    }
// For User input (Insert the Monday date of the Week to get the overview of the whole week)
    System.out.print("");
    System.out.print("Enter the date (MM-dd-yyyy): " );
      String inputDateStr = s.next();
        
    try {     
     LocalDate inputDate = LocalDate.parse(inputDateStr,DateTimeFormatter.ofPattern("MM-dd-yyyy"));
        Map<String, Map<String,Double>> weeklyHours = getWorkHoursPerWeek(filePath, inputDate);
        displayEmployeeWorkHours(employeeData, weeklyHours, inputDate.with(DayOfWeek.MONDAY), inputDate.with(DayOfWeek.FRIDAY));
    } catch (DateTimeParseException e) {
        System.out.println("Invalid date format. Please use MM-dd-yyyy");
    }
}
    
    // Method to get employee details from the Database
    public static Map<String, String[]> getEmployeeDetails(String filePath, int employeeNumberInput) {
        Map<String,String[]> employeeData = new LinkedHashMap<>();
        try( FileInputStream fis = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(fis)) {


            Sheet sheet = workbook.getSheet("Employee Details");

                if (sheet == null) {
                    System.out.println("Sheet Not Found!");
                        return employeeData;
    }

                for (Row row: sheet) {
                    if(row.getRowNum()== 0) continue; // just skipping the header
                
            Cell employeeNumberCell = row.getCell(0);
                if (employeeNumberCell == null) continue; // Skip if no Employee Number
                    
                int employeeNumber = (int)employeeNumberCell.getNumericCellValue();
                    if (employeeNumber != employeeNumberInput) continue;
                
                    
        // Relevant cells
            Cell employeeFirstNameCell = row.getCell(2);
            Cell employeeLastNameCell = row.getCell(1);
            Cell employeeBirthdayCell = row.getCell(3);
            Cell employeePositionCell = row.getCell(11);
            Cell employeeStatusCell = row.getCell(10); 
            Cell basicSalaryCell = row.getCell(13);
            Cell hourlyRateCell = row.getCell(18);
                
            
        //Extract Employee Details
            String employeeFirstName = employeeFirstNameCell.getStringCellValue();
            String employeeLastName = employeeLastNameCell.getStringCellValue();
            String employeeBirthday = formatDateCell(employeeBirthdayCell);
            String employeePosition = employeePositionCell.getStringCellValue();
            String employeeStatus = employeeStatusCell.getStringCellValue();
            int basicSalary = (int)basicSalaryCell.getNumericCellValue();
            int hourlyRate = (int)hourlyRateCell.getNumericCellValue();

        // Displaying Employee Details  
            System.out.println();
            System.out.println("===================================================");
            System.out.println("       Welcome to Motor PH Payroll System!         ");
            System.out.println("---------------------------------------------------");
            System.out.println("================ EMPLOYEE DETAILS =================");
            System.out.println("Employee Number      : " + employeeNumber);
            System.out.println("Employee Name        : " + employeeFirstName + " " + employeeLastName);
            System.out.println("Birthday             : " + employeeBirthday);
            System.out.println("---------------------------------------------------");
            System.out.println("Position             : " + employeePosition);
            System.out.println("Status               : " + employeeStatus);
            System.out.println("Hourly Rate          : ₱ " + hourlyRate);
            System.out.println("Basic Salary         : ₱ " + basicSalary);
            System.out.println();

        // Storing data for later use on the attendance records
            employeeData.put(String.valueOf(employeeNumber), new String[] {
            employeeFirstName,
            employeeLastName,
            employeeBirthday,
            employeePosition,
            employeeStatus,
            String.valueOf(basicSalary),
            String.valueOf(hourlyRate)
        });
      
    }
        }   catch (IOException e) {
                System.out.println("Error Reading File: " + e.getMessage());
            }
              return employeeData;
    }
        // Method to formate the date cells 
            private static String formatDateCell (Cell cell) {
                if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                SimpleDateFormat sdf = new SimpleDateFormat ("MM/dd/yyyy"); // Format the date
                return sdf.format(cell.getDateCellValue());
            } else if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue(); //Return as a string if not numeric
            } else { 
                return "Unknown Date Format";
            }
        }
        
        // Inserting the Database for the Attendance Record
         public static Map<String,Map<String, Double>> getWorkHoursPerWeek(String filePath, LocalDate inputDate) {
            Map<String,Map<String, Double>> weeklyHours = new HashMap<>();
            try(FileInputStream fis = new FileInputStream(new File(filePath));
                Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheet("Attendance Record");
                    if (sheet == null) {
                        System.out.println("Attendance Record Sheet not found!");
                        return weeklyHours;
                    }

        // Calculates the Start (Monday) to End (Friday)
            LocalDate weekStart = inputDate.with(DayOfWeek.MONDAY);
            LocalDate weekEnd = inputDate.with(DayOfWeek.FRIDAY);
            
            DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("MM-dd-yyyy",
            Locale.ENGLISH);
            

                for(Row row: sheet) {
                    if(row.getRowNum()== 0) continue; // Skip Header Row
                
                Cell employeeNumberCell = row.getCell(0);
                Cell dateCell = row.getCell(3);
                Cell logInCell = row.getCell (4);
                Cell logOutCell = row.getCell(5);
                
            if (employeeNumberCell == null || dateCell == null || logInCell == null || logOutCell == null) {
                continue;
            }

            String employeeNumber = getCellValueAsString(employeeNumberCell);
            LocalDate attendanceDate = getCellDateValue(dateCell);
            
        //Only process if date is present in the database
            if(attendanceDate.isBefore(weekStart)|| attendanceDate.isAfter(weekEnd)) continue;
            
            
                double hoursWorked = calculatedWorkHours(logInCell, logOutCell);
                String dayLabel = attendanceDate.format(dateFormatter); 
                
                weeklyHours.putIfAbsent(employeeNumber, new LinkedHashMap<>());
                weeklyHours.get(employeeNumber).merge(dayLabel, hoursWorked, Double::sum);
            }
        }
            catch(IOException e) {
                System.out.println("Error Reading Attendance Record: " + e. getMessage());
            }
            return weeklyHours;
        }

         public static void displayEmployeeWorkHours(Map<String, String[]> employeeData,
            Map <String,Map<String,Double>> weeklyHours, LocalDate weekStart, LocalDate weekEnd) {
                for (String employeeNumber: employeeData.keySet()) {
                    if ( weeklyHours.containsKey(employeeNumber)) {
                       
        // Diplaying the Attendance Record 
                System.out.println();
                System.out.println("================ ATTENDANCE RECORD ================");
                System.out.println();
                System.out.println("  Weekly Attendance for " + weekStart + " to " + weekEnd);
                System.out.println();
                    
                        double totalHours = 0;
                        double totalOvertime = 0;
                        int totalLateDays = 0;
                        long totalLateMinutes = 0;
                        
     
            for(Map.Entry<String, Double> entry : weeklyHours.get(employeeNumber).entrySet()) {
                String dayLabel = entry.getKey();
                double workedHours = entry.getValue();
            
                System.out.println("===================================================");
                System.out.println("Date: " + dayLabel);
                System.out.println("---------------------------------------------------");
                System.out.println("Total Worked Hours: " + String.format("%.2f", workedHours));
                            
        // Check login time
            LocalDate date = LocalDate.parse(dayLabel, DateTimeFormatter.ofPattern("MM-dd-yyyy"));
            LocalTime loginTime = getLoginTime(employeeNumber, date);
                if (loginTime != null) {
                    LocalTime graceTime = LocalTime.of(8, 10);
                if (loginTime.isAfter(graceTime)) {
                    long minutesLate = Duration.between(graceTime, loginTime).toMinutes();
                    totalLateDays++;
                    totalLateMinutes += minutesLate;
                
                System.out.println("Late: YES (" + minutesLate + " minutes late)");
            } else {
                System.out.println("Late: NO");
            }
            } else {
                System.out.println("Late: Time not available");
            }            
                    
                if (workedHours > 9) {
                double overtime = workedHours - 9;
                totalOvertime += overtime;
                
                System.out.println("Overtime Hours: " + String.format("%.2f", overtime));
            } else {
                System.out.println("Overtime Hours: 0.00");
            }
                System.out.println("===================================================");
                System.out.println("");
                    totalHours += workedHours;
            }

                         
        // WEEKLY SUMMARY DISPLAY 
                System.out.println();
                System.out.println("================== WEEKLY SUMMARY =================");
                System.out.println("Total Hours Worked for the Week: " + String.format("%.2f", totalHours)+ " hours");
                System.out.println("Total Overtime Hours: " + String.format("%.2f", totalOvertime) + " hours");
                System.out.println("Days Late: " + totalLateDays + " days");
                System.out.println("Total Late Hours: " + String.format("%.2f",totalLateMinutes / 60.0) + " hours");
                System.out.println("===================================================");
                System.out.println();
                
                    double hourlyRate = Double.parseDouble(employeeData.get(employeeNumber)[6]);
                    calculateAndDisplaySalary(totalHours, hourlyRate);

               } 
            }
         }

        // Salary Calculator
         private static void calculateAndDisplaySalary(double totalHours, double hourlyRate) {
            DecimalFormat df = new DecimalFormat("#.##");

            // Overtime
            double overtimeHours = (totalHours > 40) ? (totalHours - 40) : 0;
            double overtimeRate = hourlyRate * 1.5;
            double overtimePay = overtimeHours * overtimeRate;

            // Gross Salary Calculation
            double grossWeeklySalary = (totalHours * hourlyRate) + overtimePay;
            double grossMonthlySalary = grossWeeklySalary * 4;

            // Deductions
            double sssDeduction = calculateSSSDeduction(grossMonthlySalary);
            double philHealthDeduction = calculatePhilHealthDeduction(grossMonthlySalary);
            double pagIbigDeduction = calculatePagIbigDeduction(grossMonthlySalary);
            double taxDeduction = calculateTaxDeduction(grossWeeklySalary - 
                (sssDeduction / 4 + philHealthDeduction / 4 + pagIbigDeduction / 4));

            // Net Salary
            double netWeeklySalary = grossWeeklySalary - 
                (sssDeduction / 4 + philHealthDeduction / 4 + pagIbigDeduction / 4) - taxDeduction;

            // Display
            System.out.println("================== SALARY SUMMARY =================");
            System.out.println("Overtime Pay              : ₱ " + df.format(overtimePay));
            System.out.println("-------------------------------------");
            System.out.println("Gross Weekly Salary       : ₱ " + df.format(grossWeeklySalary));
            System.out.println("-------------------------------------");
            System.out.println("SSS Deduction             : ₱ " + df.format(sssDeduction / 4));
            System.out.println("PhilHealth Deduction      : ₱ " + df.format(philHealthDeduction / 4));
            System.out.println("Pag-Ibig Deduction        : ₱ " + df.format(pagIbigDeduction / 4));
            System.out.println("Tax Deduction             : ₱ " + df.format(taxDeduction));
            System.out.println("-------------------------------------");
            System.out.println("Net Weekly Salary         : ₱ " + df.format(netWeeklySalary));
            System.out.println("===================================================");
        }


            
            private static String getCellValueAsString(Cell cell) {
                    if (cell.getCellType() == CellType.NUMERIC){
                        return String.valueOf((int) cell.getNumericCellValue());
                    } else {
                        return cell.getStringCellValue();
                    }
                }
                    
                private static LocalDate getCellDateValue (Cell cell) {
                    if (cell.getCellType() == CellType.NUMERIC){
                        return cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault())
                                .toLocalDate();
                    } else {
                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MMM-yyy");
                        return LocalDate.parse(cell.getStringCellValue(),formatter);
                    }
                }
                
                private static LocalTime getTimeAsLocalTime(Cell cell) {
                try {
                    // Check if the cell contains a date/time value
                    if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        // Convert the date to LocalTime (just using the time part)
                        return LocalTime.of(date.getHours(), date.getMinutes(), date.getSeconds());
                    } else {
                        throw new IllegalArgumentException("Cell does not contain a valid time.");
                    }
                } catch (Exception e) {
                    System.out.println("Error parsing cell time: " + e.getMessage());
                    return LocalTime.MIN; // Return the minimum time if parsing fails
                }
            }
                
                private static double calculatedWorkHours(Cell logInCell, Cell logOutCell) {
                    try {
     
                        LocalTime loginTime = getTimeAsLocalTime(logInCell);
                        LocalTime logoutTime = getTimeAsLocalTime(logOutCell);
                        
                        double hoursWorked = (logoutTime.toSecondOfDay() - loginTime.toSecondOfDay()) / 3600.0;
                        
                    // Calculate overtime if logout is after official working hours
                          double overtimeHours = 0;
                          if (logoutTime.isAfter(LocalTime.of(17, 0))) {
                              Duration overtimeDuration = Duration.between(LocalTime.of(17, 0), logoutTime);
                              overtimeHours = overtimeDuration.toMinutes() / 60.0;
                          }

                          return hoursWorked + overtimeHours;

                        } catch (Exception e) {
                            System.out.println("Error pasing work hours." + e.getMessage());
                            return 0;
                        }
                    }
                
                private static double getTimeAsNumeric (Cell cell) {
                 try {
                    if (cell .getCellType()== CellType.NUMERIC){
                        return cell.getNumericCellValue();
                    } else {
                        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
                        Date date = sdf.parse(cell.getStringCellValue());
                        Calendar cal = Calendar.getInstance();
                        cal.setTime(date);
                        return(cal.get(Calendar.HOUR_OF_DAY) + cal.get(Calendar.MINUTE) / 60.0 +  cal.get(Calendar.SECOND)/ 3600.0) / 24.0;
                    }
                } catch (Exception e ) {
                     System.out.println("Error pasing time :"  + e.getMessage());
                     return 0;
        }
    }
            private static LocalTime getLoginTime(String employeeNumber, LocalDate date) {
                String filePath = "src//MotorPH Employee Data .xlsx";
                try (FileInputStream fis = new FileInputStream(new File(filePath));
                     Workbook workbook = new XSSFWorkbook(fis)) {

                    Sheet sheet = workbook.getSheet("Attendance Record");
                    if (sheet == null) {
                        return null;
                    }

                    for (Row row : sheet) {
                        if (row.getRowNum() == 0) continue;

                        Cell empNumberCell = row.getCell(0);
                        Cell dateCell = row.getCell(3);
                        Cell loginCell = row.getCell(4);

                        if (empNumberCell == null || dateCell == null || loginCell == null) continue;

                        String empNum = getCellValueAsString(empNumberCell);
                        LocalDate attendanceDate = getCellDateValue(dateCell);

                        if (empNum.equals(employeeNumber) && attendanceDate.equals(date)) {
                            return getTimeAsLocalTime(loginCell);
                        }
                    }
                } catch (IOException e) {
                    System.out.println("Error reading login time: " + e.getMessage());
                }

                return null;
            }

            private static double calculateSSSDeduction (double grossMonthlySalary) { 
                if (grossMonthlySalary < 3250) return 135.00;
                if (grossMonthlySalary <= 3750) return 157.50;
                if (grossMonthlySalary <= 4250) return 180.00;
                if (grossMonthlySalary <= 5250) return 225.00;
                if (grossMonthlySalary <= 5750) return 247.50;
                if (grossMonthlySalary <= 6250) return 270.00;
                if (grossMonthlySalary <= 6750) return 292.50;
                if (grossMonthlySalary <= 7250) return 315.00;
                if (grossMonthlySalary <= 7750) return 337.50;
                if (grossMonthlySalary <= 8250) return 360.00;
                if (grossMonthlySalary <= 8750) return 382.50;
                if (grossMonthlySalary <= 9250) return 405.00;
                if (grossMonthlySalary <= 9750) return 427.50;
                if (grossMonthlySalary <= 10250) return 450.00;
                if (grossMonthlySalary <= 10750) return 472.50;
                if (grossMonthlySalary <= 11250) return 495.00;
                if (grossMonthlySalary <= 11750) return 517.50;
                if (grossMonthlySalary <= 12250) return 540.00;
                if (grossMonthlySalary <= 12750) return 562.50;
                if (grossMonthlySalary <= 13250) return 585.00;
                if (grossMonthlySalary <= 13750) return 607.50;
                if (grossMonthlySalary <= 14250) return 630.00;
                if (grossMonthlySalary <= 14750) return 652.50;
                if (grossMonthlySalary <= 15250) return 675.00;
                if (grossMonthlySalary <= 15750) return 697.50;
                if (grossMonthlySalary <= 16250) return 720.00;
                if (grossMonthlySalary <= 16750) return 742.50;
                if (grossMonthlySalary <= 17250) return 765.00;
                if (grossMonthlySalary <= 17750) return 787.50;
                if (grossMonthlySalary <= 18250) return 810.00;
                if (grossMonthlySalary <= 18750) return 832.50;
                if (grossMonthlySalary <= 19250) return 855.00;
                if (grossMonthlySalary <= 19750) return 877.50;
                if (grossMonthlySalary <= 20250) return 900.00;
                if (grossMonthlySalary <= 20750) return 922.50;
                if (grossMonthlySalary <= 21250) return 945.00;
                if (grossMonthlySalary <= 21750) return 967.50;
                if (grossMonthlySalary <= 22250) return 990.00;
                if (grossMonthlySalary <= 22750) return 1012.50;
                if (grossMonthlySalary <= 23250) return 1035.00;
                if (grossMonthlySalary <= 23750) return 1057.50;
                if (grossMonthlySalary <= 24250) return 1080.00;
                if (grossMonthlySalary <= 24750) return 1102.50;
                return 1125.00;
            }

        private static double calculatePhilHealthDeduction (double grossMonthlySalary) {
         if (grossMonthlySalary == 10000) {
                    return 300;
                } else if (grossMonthlySalary >= 59999.99) {
                    double philHealthDeduction = grossMonthlySalary * 0.03;
                    return Math.min(philHealthDeduction, 1800);
                } else {
                    return 1800;
                }
            }
        private static double calculatePagIbigDeduction(double grossMonthlySalary) {
                if (grossMonthlySalary >= 1000 && grossMonthlySalary <= 1500) {
                    return grossMonthlySalary * 0.01;
                } else if (grossMonthlySalary > 1500) {
                    return grossMonthlySalary * 0.02;
                }
                return 0;
            }

            private static double calculateTaxDeduction(double taxableSalary) {
                if (taxableSalary <= 20832) return 0;
                if (taxableSalary <= 33333) return (taxableSalary - 20832) * 0.20;
                if (taxableSalary <= 66667) return (taxableSalary - 33333) * 0.25 + 2500;
                if (taxableSalary <= 166667) return (taxableSalary - 66667) * 0.30 + 10833;
                if (taxableSalary <= 666667) return (taxableSalary - 166667) * 0.32 + 40833.33;
                return (taxableSalary - 166667) * 0.35 + 200833.33;
            }
}
        









  
 



            

       
