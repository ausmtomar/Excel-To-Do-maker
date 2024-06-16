# Excel-To-Do-maker
## Subject To-Do List Organizer with Date Tracking

This Python project, tailored for your specific needs, transforms into a subject-based to-do list organizer that leverages existing information in your Excel spreadsheets. Here's a breakdown of its functionality:

**Functionality:**

1. **Data Structure:**
   - Each sheet in your Excel workbook represents a different subject.
   - Row names within each sheet correspond to specific topics within that subject.
   - Column headings (ideally the first column) represent actions like "reading," "revising," "solving dpps," etc.
   - Cell values contain dates or date formulas (e.g., `=DATE(YEAR, MONTH, DAY)`) indicating when you need to perform those actions for the respective topics.

2. **Finding To-Do Items:**
   - You'll use the provided `find_date_cells` function with a specific date (as a Python `date` object or a formula string).
   - The function searches through all subject sheets (each sheet name) and identifies cells containing that date.
   - For each matching cell, it retrieves the corresponding subject name, topic name, and action (column heading).

3. **Output:**
   - The program will display a formatted message showing the date you provided.
   - It then presents the list of to-do items for that date, categorized by subject and topic. This breakdown helps you prioritize tasks efficiently.

**Example Usage:**

Imagine you want to see your to-do list for tomorrow. You'd provide the code with tomorrow's date (or a formula to calculate it). The program would then scan your Excel file (assuming it follows the described structure) and present a list like this:

```
To-Do List for 2024-06-17 (Tomorrow)

**Subject 1:**
  * Topic A: Reading (Column Heading for Action)
  * Topic B: Solving DPPs (Column Heading for Action)

**Subject 2:**
  * Topic C: Revising (Column Heading for Action)
```

**Benefits:**

- **Organized To-Do List:** This project helps you visualize your workload for a specific date, categorized by subjects and topics, promoting better organization and task management.
- **Date Tracking:** By utilizing dates or date formulas in your Excel file, you can easily track upcoming deadlines and stay on top of your academic responsibilities.
- **Flexibility:** The code works with existing information in your Excel file, so you don't need to create a separate to-do list application.

**Additional Considerations:**

- Remember to modify the code's assumptions about row and column headings if your Excel files are structured differently.
- If your Excel file doesn't currently use date formulas, you'll need to update it manually before using this script.
- The code currently doesn't handle editing or adding to-do items. You'd need to modify it or use the Excel file itself for those tasks.
