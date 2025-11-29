# Configuration for the BaiHelper module
$Script:Config = @{
    # Base URLs
    BaseUrl = "https://nasptournaments.org"
    LoginPath = "/userutilities/login.aspx"
    ScoreSheetPath = "/Schoolmgr/SeasonScoreSheet.aspx"
    
    # ASP.NET Control IDs
    Controls = @{
        Username = "ctl00`$ContentPlaceHolder1`$TextBox_username"
        Password = "ctl00`$ContentPlaceHolder1`$TextBox_password"
        ExportButton = "ctl00`$ContentPlaceHolder1`$Button_export"
        ReturnToSchoolButton = "ctl00`$ContentPlaceHolder1`$Button_return_school"
        SchoolLabel = "ctl00_ContentPlaceHolder1_Label_school_name"
    }
    
    # Regex Patterns for HTML parsing
    Patterns = @{
        SchoolName = '<span[^>]*id="ctl00_ContentPlaceHolder1_Label_school_name"[^>]*>([^<]+)</span>'
        SeasonDropdown1 = '<select[^>]*id="([^"]*Season[^"]*)"'
        SeasonDropdown2 = '<select[^>]*id="([^"]*ddl[^"]*)"'
        DefaultSeason1 = '<option[^>]*selected[^>]*>([^<]+)</option>'
        DefaultSeason2 = '<select[^>]*id="[^"]*Season[^"]*"[^>]*>[\s\S]*?<option[^>]*value="[^"]*"[^>]*>([^<]+)</option>'
        DataTable = '<table[^>]*(?:class="[^"]*(?:grid|data|score)[^"]*"|id="[^"]*(?:gv|grid|tbl)[^"]*")[^>]*>([\s\S]*?)</table>'
        TableHeader = '<th[^>]*>([\s\S]*?)</th>'
        TableRow = '<tr[^>]*>([\s\S]*?)</tr>'
        TableCell = '<td[^>]*>([\s\S]*?)</td>'
        CsvContent = '^[\w\s,"]+\r?\n'
    }
    
    # File and path settings
    FileSettings = @{
        InvalidPathChars = '[<>:"/\\|?*\[\]]'
        TimestampFormat = "yyyyMMdd_HHmmss"
        Encoding = "UTF8"
    }
}