import win32com.client as win32

def create_gantt_chart():
    # Create an instance of the Microsoft Project application
    app = win32.Dispatch("C:\Program Files\Microsoft Office\Office15\WINPROJ.Application")

    # Create a new project
    project = app.NewProject

    # Set the start date of the project
    project.StartDate = "09/06/2023"  # Format: MM/DD/YYYY

    # Add tasks to the project
    task1 = project.Tasks.Add("Task 1")
    task1.Start = "09/06/2023"
    task1.Duration = "5d"  # 5 days

    task2 = project.Tasks.Add("Task 2")
    task2.Start = "09/11/2023"
    task2.Duration = "3d"  # 3 days

    # Create a Gantt chart view
    view = project.Views.Add("Gantt Chart", view_type=1)  # 1 represents Gantt chart view

    # Display the Gantt chart view
    view.Apply()

    # Save the project as a file
    project.SaveAs("GanttChart.mpp")

    # Close the project and the application
    project.Close()
    app.Quit()

# Call the function to create the Gantt chart
create_gantt_chart()
