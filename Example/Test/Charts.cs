using RapidExcel.Attributes;

namespace Example.Test;

public class Charts
{
    [ExcelColumn("chart_position")]
    public int Rank { get; set; }
    [ExcelColumn("chart_date")]
    public DateTime DateOfChart { get; set; }
    [ExcelColumn("chart_song")]
    public string Song { get; set; } = null!;
    [ExcelColumn("performer")]
    public string Performer { get; set; } = null!;
    [ExcelColumn("time_on_chart")]
    public int TimesOnChart { get; set; }
}
