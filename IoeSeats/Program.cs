using ExcelDataReader;
using System.Diagnostics;

int college = 0; // 0 for pulchowk, 1 for wrc, 2 for erc and 3 for thapathali, 4 chitwan
string filePath = "C:\\Users\\blackjack\\Downloads\\BESejalAll.xlsx";

Dictionary<int, List<int>> seatsIOE = new Dictionary<int, List<int>>()
{
    {1,  new List<int> {108, 36,  36,  36,  0}}, // civil regular 
    {2,  new List<int> {84,  108, 108, 108, 0}},// civil fullfee

    {3,  new List<int> {24,  0,   12,  12,  6 }}, // architecture regular
    {4,  new List<int> {24,  0,   36,  36,  18}}, // architecture fullfee

    {5,  new List<int> {36,  12,  12,  0,   0}}, // electrical fullfee
    {6,  new List<int> {60,  36,  36,  0,   0}}, // electrical fullfee

    {7,  new List<int> {24,  12,  12,  12,  0} }, // electronics and computer regular
    {8,  new List<int> {24,  36,  36,  36,  0}}, // electronics and computer fullfee

    {9,  new List<int> {24,  12,  24,  12,  0}}, // mechanical regular
    {10, new List<int> {24,  36,  72,  36,  0}},// mechanical fullfee 

    {11, new List<int> {36,  12,  24,  12,  0}},// computer regular
    {12, new List<int> {60,  36,  72,  36,  0}},// computer fullfee

    {13, new List<int> {0,   0,   12,  0,   0}},// agriculture regular
    {14, new List<int> {0,   0,   36,  0,   0}},// agriculture fullfee

    {15, new List<int> {0,   0,   0,   12,  0}},// industrial regular
    {16, new List<int> {0,   0,   0,   36,  0}},// industrial fullfee

    {17, new List<int> {0,   12,  0,   0,  0}},// geomatics regular
    {18, new List<int> {0,   36,  0,   0,  0}},// geomatics fullfee

    {19, new List<int> {0,   12,  0,   12,  0}},// automobile regular
    {20, new List<int> {0,   36,  0,   36,  0}},// automobile fullfee

    {21, new List<int> {0,   0,   0,   0,  0}},//   regular
    {22, new List<int> {0,   0,   0,   0,  0}},//   fullfee

    {23, new List<int> {0,   0,   0,   0,  0}},//   regular
    {24, new List<int> {0,   0,   0,   0,  0}},//   fullfee

    {25, new List<int> {0,   0,   0,   0,  0}},//   regular
    {26, new List<int> {0,   0,   0,   0,  0}},//   fullfee

    {27, new List<int> {12,  0,   0,   0,  0}},// aerospace regular
    {28, new List<int> {36,  0,   0,   0,  0}},// aerospace fullfee

    {29, new List<int> {12,  0,   0,   0,  0}},// chemical regular
    {30, new List<int> {36,  0,   0,   0,  0}},// chemical fullfee
};

Dictionary<int, List<int>> finalAllotedSeats = new Dictionary<int, List<int>>();
foreach (var item in seatsIOE)
{
    finalAllotedSeats.Add(item.Key, new List<int>());
}

var vmList = new List<RankViewModel>();

//should be in format of
// 1.sn, 2.rollno, 3.name, 4.gender, 5.location, 6.rank , 7priority, 18remarks
// repsent gender by "Male", "Female"
int collegeSheet = 0;
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
{
    using (var reader = ExcelReaderFactory.CreateReader(stream))
    {
        do
        {
            while (reader.Read()) //Each ROW
            {
                var vm = new RankViewModel()
                {
                    CollegeNumber = collegeSheet,
                };

                for (int column = 0; column < reader.FieldCount; column++)
                {
                    var value = reader.GetValue(column);

                    Trace.WriteLine(value);

                    string valueString = value?.ToString() ?? string.Empty;
                    if (string.IsNullOrEmpty(valueString))
                        continue;


                    switch (column)
                    {
                        case 0:
                            Int32.TryParse(valueString, out int v1);
                            vm.SN = v1;
                            break;
                        case 1:
                            Int32.TryParse(valueString, out int v2);
                            vm.RollNo = v2;
                            break;
                        case 2:
                            vm.Name = valueString;
                            break;
                        case 3:
                            vm.Gender = valueString;
                            break;
                        case 4:
                            vm.Location = valueString;
                            break;
                        case 5:
                            Int32.TryParse(valueString, out int v3);
                            vm.Rank = v3;
                            break;
                        case 6:
                        case 7:
                        case 8:
                        case 9:
                        case 10:
                        case 11:
                        case 12:
                        case 13:
                        case 14:
                        case 15:
                        case 16:
                        case 17:
                        case 18:
                            Int32.TryParse(valueString, out int v4);
                            vm.Priority.Add(v4);
                            break;
                        case 19:
                            vm.Remarks = valueString;
                            break;
                        default:
                            break;
                    }
                }

                vmList.Add(vm);
            }
            collegeSheet++;
        } while (reader.NextResult()); //Move to NEXT SHEET
    }
}

var group = vmList.GroupBy(n => n.CollegeNumber).Select(n => new { MetricName = n.Key, MetricCount = n.Count() }).OrderBy(n => n.MetricName).ToList();

var validRows = vmList.Where(x => x.Rank != 0).OrderBy(x => x.Rank).ToList();
var numberOfSheets = validRows.Select(x => x.CollegeNumber).Distinct().OrderBy(x=>x).ToList();
foreach (var c in numberOfSheets)
{
    var finalList = validRows.Where(x => x.CollegeNumber == c).ToList();

    foreach (var ind in finalList)
    {
        var priority = ind.Priority; //.OrderByDescending(x => x).ToList();
        bool isMale = ind.Gender == "Male" ? true : false;
        foreach (var p in priority)
        {
            if (p == 0) continue;
            var seatsValue = seatsIOE[p][college];
            var finalAlloted = finalAllotedSeats[p];

            decimal abs = Math.Abs(seatsValue);
            var quotaString = (abs / 10).ToString().Split('.').FirstOrDefault();
            Int32.TryParse(quotaString, out int quota);

            if ((!isMale && finalAlloted.Count < seatsValue)
                || (isMale && finalAlloted.Count < seatsValue - quota))
            {
                finalAlloted.Add(ind.Rank);// rank added
                break;
            }
        }
    }

    Trace.WriteLine($"college- {c}");
    foreach (var item in finalAllotedSeats)
    {
        string value = string.Empty;
        foreach (var inner in item.Value)
        {
            var ind = finalList.FirstOrDefault(x => x.Rank == inner);
            char gender = ind?.Gender == "Male" ? 'M' : 'F';
            value = value + inner + "(" + gender + ")" + ',';
        }

        Trace.WriteLine(item.Key + "--->" + value);
    }
    Trace.WriteLine($"completed");

    finalAllotedSeats.ToList().ForEach(x => x.Value.Clear());
}



Trace.WriteLine("completed");

public class RankViewModel
{
    public int CollegeNumber { get; set; }
    public int SN { get; set; }
    public int RollNo { get; set; }
    public string Name { get; set; } = String.Empty;
    public string Gender { get; set; } = String.Empty;
    public string Location { get; set; } = String.Empty;
    public int Rank { get; set; }
    public List<int> Priority { get; set; } = new List<int>();
    public string Remarks { get; set; } = String.Empty;
}


