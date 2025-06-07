📊 Student Marks Management System (Python + Excel)
This is a simple Student Marks Management System built using Python and openpyxl to manage student records in an Excel file. It allows you to add, update, and display student marks with automatic calculation of total, percentage, grade, and pass/fail status.

📁 Features
✅ Create or open an Excel file
✅ Add new student records
✅ Update existing student marks
✅ Auto-calculate total, percentage, grade, and pass/fail
✅ Display all students in the console
✅ User-friendly CLI interface

🛠️ Technologies Used
Python 3

openpyxl (for Excel operations)

🧮 Grading Criteria
Percentage	Grade
90+	A+
80–89	A
70–79	B
60–69	C
50–59	D
Below 50	F

▶️ How to Run

Install dependencies: pip install openpyxl

▶️Run the script:

python Excel_student_marks_sys.py
Follow the CLI options to add, update, or view student records.

📂 Output

A file named Student_mark_sys.xlsx will be created in the same directory.

The first row is the title.
The second row contains column headers like Roll No, Name, Marks, Grade, and Pass/Fail.
![image](https://github.com/user-attachments/assets/4f7c8844-b0a0-4458-8f65-54b9c02885cc)

📸 Sample Menu Output

Welcome to the Student Marks Management System

1. Add Student  
2. Update Student Marks  
3. Display All Students  
4. Exit
   
🧪 Example Entry

Roll No	Name	Maths	OS	English	DB	Prog	Total	%	Grade	Result
101	Alice	78	84	88	90	85	425	85.0	A	Pass

❗ Notes

Make sure Student_mark_sys.xlsx is not open in Excel while updating.
Marks must be between 0 and 100.
Unique roll numbers are enforced.

📬 Contact
For any questions or improvements, feel free to raise an issue or pull request.
