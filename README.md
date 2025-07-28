🟩 1. Excel Process Scope
This is the main container where all Excel-related activities are performed.

🟦 2. Use Excel File
File Path: Points to the Patient_Details Excel file.

The file is opened and accessed for reading and writing.

Save changes is checked, meaning any changes made to the file will be saved.

🟨 3. Read Range ("Sheet1")
Reads all the data from Sheet1 of the Excel file into a DataTable (dt).

🔁 4. For Each Row in Data Table
Iterates through each row (CurrentRow) in the DataTable.

Inside the loop:
🟩 5. Assign Activities
tomorrow = Now.AddDays(1).ToString
➤ Gets the date for tomorrow.

dtFiltered = (From r In dt.AsEnumerable() ...)
➤ Seems like some filtering is being prepared but unused here (possibly remnant/placeholder).

🟥 6. Condition (Check if ID is not empty)
vb
Copy
Edit
If Not String.IsNullOrWhiteSpace(CurrentRow("ID").ToString)
If ID exists (i.e., the row contains data):

🔷 7. Assign Values from CurrentRow
Extracts appointment-related details:

Email, Name, apptDate, apptTime, manualDate, systemDate

difference = Math.Round(DateTime.Parse(manualDate).Subtract(systemDate).TotalDays)

Calculates how many days between the appointment and today.

⚠ 8. Condition: If Difference = 1
Means appointment is tomorrow.

📧 9. Send Outlook Mail Message
Sends a reminder email to the patient.

Subject: “Appointment Reminder”

Body: Personalized reminder with patient name, date, and time.

✅ 10. Mark Status as "Sent"
CurrentRow("Status") = "Sent"

Writes the updated row back to the Excel file (Write Range).

❌ 11. Else – If Not Sending
If the difference is not 1:

Status = "unSent"

Writes this back to Excel.

📦 12. Final Else – Fallback
If none of the above executes, it shows a Message Box saying "mail sent" (though placement suggests it may just be a final step).

📝 Summary of Logic:
Reads patient appointment data.

Checks if the appointment is tomorrow.

If yes → sends an email reminder → marks row as "Sent".

If no → marks as "unSent".

Saves all changes back to Excel.

✅ Use Case:
Perfect for clinic/hospital automation, helping:

Send reminders automatically.

Keep a status log for each patient.
