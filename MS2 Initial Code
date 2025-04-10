import java.util.Scanner;

public class EmployeeData {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        int[] employeeNumbers = {10001, 10002, 10003, 10004, 10005, 10006, 10007, 10008, 10009, 10010, 
                                 10011, 10012, 10013, 10014, 10015, 10016, 10017, 10018, 10019, 10020, 
                                 10021, 10022, 10023, 10024, 10025, 10026, 10027, 10028, 10029, 10030, 
                                 10031, 10032, 10033, 10034};

        String[] basicSalaries = {"90000", "60000", "60000", "60000", "52670", "52670", "42975", "22500", "22500", 
                                  "52670", "50825", "38475", "24000", "24000", "53500", "42975", "41850", "22500", 
                                  "22500", "23250", "23250", "24000", "22500", "22500", "24000", "24750", "24750", 
                                  "24000", "22500", "22500", "22500", "52670", "52670", "52670"};

        // Prompt for employee number
        System.out.print("Enter employee number: ");
        int employeeNumber = scanner.nextInt();

        // Search for employee index
        int index = -1;
        for (int i = 0; i < employeeNumbers.length; i++) {
            if (employeeNumbers[i] == employeeNumber) {
                index = i;
                break;
            }
        }

        if (index == -1) {
            System.out.println("Employee number not found.");
        } else {
            double basicSalary = Double.parseDouble(basicSalaries[index]);

            // Prompt for hours worked
            System.out.print("Enter hours worked this week: ");
            double hoursWorked = scanner.nextDouble();

            // Calculate total earnings
            double hourlyRate = basicSalary / 160;
            double totalEarnings = hoursWorked * hourlyRate;

            // Calculate deductions
            double sss = calculateSSS(basicSalary);
            double philHealth = calculatePhilHealth(basicSalary);
            double pagIbig = calculatePagIbig(basicSalary);
            double withholdingTax = calculateWithholdingTax(totalEarnings);

            // Total salary deductions
            double salaryDeductions = sss + philHealth + pagIbig;

            // Calculate net income
            double netIncome = (totalEarnings - salaryDeductions) - withholdingTax;

            // Display results
            System.out.println("\nEmployee #" + employeeNumber);
            System.out.println("Basic Salary: " + basicSalary + " Php");
            System.out.println("Hourly Rate: " + hourlyRate + " Php");
            System.out.println("Hours Worked: " + hoursWorked + " hrs");
            System.out.println("Total Earnings: " + totalEarnings + " Php");
            System.out.println("\nDeductions:");
            System.out.println("SSS: " + sss + " Php");
            System.out.println("PhilHealth: " + philHealth + " Php");
            System.out.println("Pag-IBIG: " + pagIbig + " Php");
            System.out.println("Withholding Tax (on total earnings): " + withholdingTax + " Php");
            System.out.println("\nSalary Deductions: " + salaryDeductions + " Php");
            System.out.println("Net Income: " + netIncome + " Php");
        }

        scanner.close();
    }

    public static double calculateSSS(double salary) {
        if (salary >= 22250 && salary < 22750) return 1012.50;
        if (salary >= 22750 && salary < 23250) return 1035.00;
        if (salary >= 23250 && salary < 23750) return 1057.50;
        if (salary >= 23750 && salary < 24250) return 1080.00;
        if (salary >= 24250 && salary < 24750) return 1102.50;
        if (salary >= 24750) return 1125.00;
        return 0.0; // No deduction if salary is below the lowest range
    }

    public static double calculatePhilHealth(double salary) {
        if (salary >= 10000.01 && salary <= 59999.99) return salary * 0.03;
        if (salary >= 60000) return salary * 0.03;
        return 0.0;
    }

    public static double calculatePagIbig(double salary) {
        double deduction = salary * 0.02;
        return (deduction > 100) ? 100 : deduction; // Max deduction is 100
    }

    public static double calculateWithholdingTax(double totalEarnings) {
        if (totalEarnings >= 20833 && totalEarnings < 33333) return (totalEarnings - 20833) * 0.20;
        if (totalEarnings >= 33333 && totalEarnings < 66667) return 2500 + (totalEarnings - 33333) * 0.25;
        if (totalEarnings >= 66667 && totalEarnings < 166667) return 10833 + (totalEarnings - 66667) * 0.30;
        if (totalEarnings >= 166667 && totalEarnings < 666667) return 40833.33 + (totalEarnings - 166667) * 0.32;
        if (totalEarnings >= 666667) return 200833.33 + (totalEarnings - 666667) * 0.35;
        return 0.0; // No tax if below 20,833
    }
}
