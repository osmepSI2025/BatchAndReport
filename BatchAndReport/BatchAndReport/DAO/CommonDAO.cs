using System;
using System.Globalization;
using System.Text;

namespace BatchAndReport.DAO
{
    public class CommonDAO
    {
        public static string[] ToThaiDateString(DateTime date)
        {
            int thaiYear = date.Year + 543;
            string thaiMonth = date.ToString("MMMM", new CultureInfo("th-TH"));
            string day = ToThaiNumber(date.Day);
            string year = ToThaiNumber(thaiYear);
            return new string[] { day, thaiMonth, year };
        }

        public static string ToThaiDateStringCovert(DateTime date)
        {
            int thaiYear = date.Year + 543;
            string thaiMonth = date.ToString("MMMM", new CultureInfo("th-TH"));
            string day = date.Day.ToString();
            if (thaiYear < 2500)
            {
                thaiYear = thaiYear + 543;
            }
            return $"วันที่ {ToThaiNumber(date.Day)} เดือน {thaiMonth} พ.ศ. {ToThaiNumber(thaiYear)}";
        }
        public static string ToThaiDateStringCovert_Noname(DateTime date)
        {
            int thaiYear = date.Year + 543;
            string thaiMonth = date.ToString("MMMM", new CultureInfo("th-TH"));
            string day = date.Day.ToString();
            if (thaiYear < 2500)
            {
                thaiYear = thaiYear + 543;
            }
            return $"วันที่ {ToThaiNumber(date.Day)} {thaiMonth}  {ToThaiNumber(thaiYear)}";
        }
        public static string ToArabicDateStringCovert(DateTime date)
        {
            int thaiYear = date.Year + 543;
            string thaiMonth = date.ToString("MMMM", new CultureInfo("th-TH"));
            string day = date.Day.ToString();
            if (thaiYear < 2500)
            {
                thaiYear = thaiYear + 543;
            }
            return $"วันที่ {date.Day} {thaiMonth} {thaiYear}";
        }

        public static string ToThaiNumber(int number)
        {
            return ToThaiNumber(number.ToString());
        }

        public static string ToThaiNumber(string number)
        {
            char[] thaiDigits = { '๐', '๑', '๒', '๓', '๔', '๕', '๖', '๗', '๘', '๙' };
            var sb = new StringBuilder();
            foreach (char c in number)
            {
                if (char.IsDigit(c))
                    sb.Append(thaiDigits[c - '0']);
                else
                    sb.Append(c);
            }
            return sb.ToString();
        }

        // Function: Convert number to Thai text (Baht/Satang)
        public static string NumberToThaiText(decimal amount)
        {
            string[] numText = { "ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า" };
            string[] rankText = { "", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน", "สิบล้าน", "ร้อยล้าน", "พันล้าน", "หมื่นล้าน", "แสนล้าน", "ล้านล้าน" };

            string bahtText = "";
            string satangText = "";

            string[] parts = amount.ToString("0.00").Split('.');
            long baht = long.Parse(parts[0]);
            int satang = int.Parse(parts[1]);

            bahtText = ConvertIntegerToThaiText(baht, numText, rankText);
            if (baht == 0) bahtText = "ศูนย์";

            if (satang > 0)
            {
                satangText = ConvertIntegerToThaiText(satang, numText, rankText) + "สตางค์";
            }
            else
            {
                satangText = "ถ้วน";
            }

            return bahtText + "บาท" + satangText;
        }

        private static string ConvertIntegerToThaiText(long number, string[] numText, string[] rankText)
        {
            StringBuilder result = new StringBuilder();
            string numStr = number.ToString();
            int len = numStr.Length;

            for (int i = 0; i < len; i++)
            {
                int digit = int.Parse(numStr[i].ToString());
                int rank = len - i - 1;

                if (digit == 0) continue;

                if (rank == 0 && digit == 1 && len > 1)
                    result.Append("เอ็ด");
                else if (rank == 1 && digit == 2)
                    result.Append("ยี่");
                else if (rank == 1 && digit == 1)
                    result.Append("");
                else
                    result.Append(numText[digit]);

                result.Append(rankText[rank]);
            }
            return result.ToString();
        }
        public static string ConvertStringArabicToThaiNumerals(string input)
        {
            // สร้าง Dictionary เพื่อเก็บคู่การจับคู่ระหว่างเลขอารบิกและเลขไทย
            var arabicToThai = new Dictionary<char, char>
    {
        {'0', '๐'},
        {'1', '๑'},
        {'2', '๒'},
        {'3', '๓'},
        {'4', '๔'},
        {'5', '๕'},
        {'6', '๖'},
        {'7', '๗'},
        {'8', '๘'},
        {'9', '๙'}
    };
            if (input != null)
            {
                // ใช้ StringBuilder เพื่อประสิทธิภาพในการจัดการสตริง
                var sb = new StringBuilder();

                foreach (char c in input)
                {
                    // ตรวจสอบว่าอักขระปัจจุบันมีใน Dictionary หรือไม่
                    if (arabicToThai.ContainsKey(c))
                    {
                        // ถ้ามี ให้เพิ่มเลขไทยลงใน StringBuilder
                        sb.Append(arabicToThai[c]);
                    }
                    else
                    {
                        // ถ้าไม่มี ให้เพิ่มอักขระเดิมลงไป
                        sb.Append(c);
                    }
                }

                return sb.ToString();
            }
            else {

                return "";
            }
        }
        public static string ConvertCurrencyToThaiNumerals(decimal amount)
        {
            // จัดรูปแบบตัวเลขให้มี comma ตามหลักสากลก่อน
            // "N0" คือการจัดรูปแบบตัวเลขจำนวนเต็มที่ไม่มีทศนิยมและมี comma คั่น
            string formattedAmount = amount.ToString("N0", CultureInfo.InvariantCulture);

            // Dictionary สำหรับแปลงเลขอารบิกเป็นเลขไทย
            var arabicToThai = new Dictionary<char, char>
    {
        {'0', '๐'}, {'1', '๑'}, {'2', '๒'}, {'3', '๓'}, {'4', '๔'},
        {'5', '๕'}, {'6', '๖'}, {'7', '๗'}, {'8', '๘'}, {'9', '๙'}
    };

            var result = new StringBuilder();
            foreach (char c in formattedAmount)
            {
                // ถ้าเป็นตัวเลข ให้แปลงเป็นเลขไทย
                if (arabicToThai.ContainsKey(c))
                {
                    result.Append(arabicToThai[c]);
                }
                // ถ้าเป็น comma หรืออักขระอื่นๆ ให้คงเดิม
                else
                {
                    result.Append(c);
                }
            }
            return result.ToString();
        }
    }
}

