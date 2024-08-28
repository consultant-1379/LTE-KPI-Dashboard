from System import DateTime
Document.Properties["KPIdashboardrefresh"] = DateTime.Now
Document.ActivePageReference = Document.Pages[1]