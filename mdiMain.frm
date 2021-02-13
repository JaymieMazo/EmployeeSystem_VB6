VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Employee Management System"
   ClientHeight    =   12300
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   17685
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuEmp1 
      Caption         =   "Employees"
      Begin VB.Menu mnuSched 
         Caption         =   "Schedules"
      End
      Begin VB.Menu mnuOT 
         Caption         =   "Overtime"
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "Breaktime Monitor"
      End
      Begin VB.Menu mnuAttendance 
         Caption         =   "Attendance"
         Begin VB.Menu mnuAbsent 
            Caption         =   "Absences"
         End
         Begin VB.Menu mnuPerfect 
            Caption         =   "Perfect Attendance"
         End
      End
      Begin VB.Menu mnuAwards 
         Caption         =   "Awards"
         Begin VB.Menu mnuEarly 
            Caption         =   "Early Bird"
         End
         Begin VB.Menu mnuYrs 
            Caption         =   "Years in Service"
         End
         Begin VB.Menu mnuProductivity 
            Caption         =   "Productivity"
         End
         Begin VB.Menu mnuPromotion 
            Caption         =   "Promotion"
         End
      End
      Begin VB.Menu mnuResignation 
         Caption         =   "Resignation"
      End
      Begin VB.Menu mnuViolation 
         Caption         =   "Violations"
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "Master Maintenance"
      Begin VB.Menu mnuUsers 
         Caption         =   "Users"
      End
      Begin VB.Menu mnuEmp 
         Caption         =   "Employees"
      End
      Begin VB.Menu mnuDept 
         Caption         =   "Departments"
      End
      Begin VB.Menu mnuSec 
         Caption         =   "Sections"
      End
      Begin VB.Menu mnuDesignation 
         Caption         =   "Designations"
      End
      Begin VB.Menu mnuBreaklist 
         Caption         =   "Breaktime"
      End
      Begin VB.Menu mnuLeave 
         Caption         =   "Leave"
      End
      Begin VB.Menu mnuLeaveType 
         Caption         =   "Types of Leave"
      End
   End
   Begin VB.Menu mnuForms 
      Caption         =   "Forms"
   End
   Begin VB.Menu mnuLogout 
      Caption         =   "Logout"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuEmp_Click()
frm_m_employees.Show
End Sub

Private Sub mnuLogout_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub mnuUsers_Click()
frm_m_user.Show
End Sub
