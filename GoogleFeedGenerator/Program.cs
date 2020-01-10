using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.Core.Objects;
using System.Data.OleDb;
using System.Xml;

namespace GoogleFeedGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            string output = "";
            string itemText = "";
            string line;
            string productId;
            string productSku;
            string productName;
            string productType;
            string categoryStr;
            string weight;
            string description;
            string price;
            string visible;
            string allowed;
            string freeShip;
            string fixedShip;
            string width;
            string height;
            string depth;
            string brand;
            string url;
            string img;
            string gtin;
            string mpn;
            string googleCategory;
            string googleEnabled;
            string dateAdded;
            string shippingPrice;
            string l0;
            string l1;
            string expDate;
            string itemTextTmp;
            long gtinCheck;
            DateTime dt = DateTime.Now;
            DateTime expDt = DateTime.Now.AddMonths(3);
            String currDate = dt.Year.ToString() + "-";
            if (dt.Month < 10) currDate = currDate + "0";
            currDate = currDate + dt.Month.ToString() + "-";
            if (dt.Day < 10) currDate = currDate + "0";
            currDate = currDate + dt.Day.ToString();
            String currTime = dt.Hour.ToString() + ":";
            if (dt.Minute < 10) currTime = currTime + "0";
            currTime = currTime + dt.Minute.ToString() + ":";
            if (dt.Second < 10) currTime = currTime + "0";
            currTime = currTime + dt.Second.ToString();

            expDate = expDt.Year.ToString() + "-";
            if (expDt.Month < 10) expDate = expDate + "0";
            expDate = expDate + expDt.Month.ToString() + "-";
            if (expDt.Day < 10) expDate = expDate + "0";
            expDate = expDate + expDt.Day.ToString();


            output = output + "<?xml version=\"1.0\"?>\r\n";
            output = output + "        <feed xmlns=\"http://www.w3.org/2005/Atom\" xmlns:g=\"http://base.google.com/ns/1.0\">\r\n";
            output = output + "            <title>Performance Line Tool Center</title>\r\n";
            output = output + "            <link rel=\"self\" href=\"https://www.performancetoolcenter.com\"/>\r\n";
            output = output + "            <updated>" + currDate + "T" + currTime + "Z</updated>\r\n";
            output = output + "            <author>\r\n";
            output = output + "                <name>Performance Line Tool Center</name>\r\n";
            output = output + "            </author>\r\n";
            output = output + "            <id>tag:" + currDate.Replace("-", "") + currTime.Replace(":", "") + "</id>\r\n";

            SqlConnection conn = new SqlConnection("Data Source=pltcsrv01;Initial Catalog=PLTC;Persist Security Info=True;User ID=AppLogin;Password=Pltc6000#");

            conn.Open();

            // Read the file and display it line by line.  
            System.IO.StreamReader file =
                new System.IO.StreamReader(@"c:\feed\Google_Export.csv");
            System.IO.StreamReader itemFile =
                new System.IO.StreamReader(@"c:\feed\item.txt");
            while ((line = itemFile.ReadLine()) != null)
            {
                itemText = itemText + line + "\r\n";
            }

            
                while ((line = file.ReadLine()) != null)
            {
                l0 = "$5000";

                line = line.Replace("\"\"", "|||***|||");

                CsvSubString css = new CsvSubString();
                css = ReadCsvItem(line);
                productId = css.Item;

                css = ReadCsvItem(css.Remainder);
                productSku = css.Item;

                css = ReadCsvItem(css.Remainder);
                productName = css.Item;

                css = ReadCsvItem(css.Remainder);
                categoryStr = css.Item;

                css = ReadCsvItem(css.Remainder);
                
                if (css.Item == "0.0000") { css.Item = "1.0000"; }
                if (css.Item != "Weight")
                {
                    if (Convert.ToDouble(css.Item) > 140) { css.Item = "140.0000"; }
                }

                weight = css.Item;

                css = ReadCsvItem(css.Remainder);
                description = css.Item;

                css = ReadCsvItem(css.Remainder);
                price = css.Item;

                if (price != "Calculated Price")
                {
                    if (Convert.ToDouble(price) < 5000) { l0 = "$1000-$4999"; }
                    if (Convert.ToDouble(price) < 1000) { l0 = "$500-$999"; }
                    if (Convert.ToDouble(price) < 500) { l0 = "$250-$499"; }
                    if (Convert.ToDouble(price) < 250) { l0 = "$100-$249"; }
                    if (Convert.ToDouble(price) < 100) { l0 = "$50-$100"; }
                    if (Convert.ToDouble(price) < 50) { l0 = "<$50"; }
                }

                css = ReadCsvItem(css.Remainder);
                visible = css.Item;

                css = ReadCsvItem(css.Remainder);
                allowed = css.Item;

                css = ReadCsvItem(css.Remainder);
                freeShip = css.Item;

                css = ReadCsvItem(css.Remainder);
                fixedShip = css.Item;

                css = ReadCsvItem(css.Remainder);
                if (css.Item == "0.0000") { css.Item = "1.0000"; }
                if (css.Item != "Width")
                {
                    if (Convert.ToDouble(css.Item) > 80) { css.Item = "80.0000"; }
                }
                width = css.Item;

                css = ReadCsvItem(css.Remainder);
                if (css.Item == "0.0000") { css.Item = "1.0000"; }
                if (css.Item != "Height")
                {
                    if (Convert.ToDouble(css.Item) > 80) { css.Item = "80.0000"; }
                }
                height = css.Item;

                css = ReadCsvItem(css.Remainder);
                if (css.Item == "0.0000") { css.Item = "1.0000"; }
                if (css.Item != "Depth")
                {
                    if (Convert.ToDouble(css.Item) > 80) { css.Item = "80.0000"; }
                }
                depth = css.Item;

                css = ReadCsvItem(css.Remainder);
                brand = css.Item;

                css = ReadCsvItem(css.Remainder);
                url = css.Item;

                css = ReadCsvItem(css.Remainder);
                img = css.Item.Replace("Product Image URL: ", "");

                string[] imgs = new string[20];

                imgs = img.Split('|');

                if (imgs[0] != null) { img = imgs[0]; }

                css = ReadCsvItem(css.Remainder);
                gtin = css.Item;

                css = ReadCsvItem(css.Remainder);
                mpn = css.Item;

                css = ReadCsvItem(css.Remainder);
                googleCategory = FindGoogleCategory(css.Item, conn);
                productType = css.Item;

                css = ReadCsvItem(css.Remainder);
                googleEnabled = css.Item;

                dateAdded = css.Remainder;

                l1 = productType;

                gtinCheck = 0;
                long.TryParse(gtin, out gtinCheck);
                if (gtinCheck == 0)
                {
                    gtin = "";
                }

                if (gtin.Length < 12 || gtin.Length > 14)
                {
                    gtin = "";
                }


                if (productName != "Product Name")
                {
                    if ((googleEnabled == "1" || Convert.ToDateTime(dateAdded) > Convert.ToDateTime("09/08/2019")) && visible == "1" && allowed == "1")
                    {
                        if (gtin == "" || mpn == "")
                        {
                            itemTextTmp = itemText.Replace("<g:gtin><![CDATA[%%GTIN%%]]></g:gtin>", "<g:identifier_exists>FALSE</g:identifier_exists>") 
                                .Replace("<g:brand><![CDATA[%%BRAND%%]]></g:brand>", "")
                                .Replace("<g:mpn><![CDATA[%%MPN%%]]></g:mpn>", "");
                            
                        }
                        else
                        {
                            itemTextTmp = itemText;
                        }
                        if (freeShip == "0" && fixedShip == "0.0000") {
                            itemTextTmp = itemTextTmp.Replace("<g:shipping><g:price><![CDATA[%%SHIPPINGPRICE%% USD]]></g:price></g:shipping>", "");
                        }

                            Console.WriteLine(productId);
                        output = output + itemTextTmp.Replace("%%PRODUCTID%%", productId).Replace("%%PRODUCTSKU%%", productSku).Replace("%%TITLE%%", productName)
                                 .Replace("%%LINK%%", url).Replace("%%DESCRIPTION%%", description).Replace("%%BRAND%%", brand).Replace("%%IMG%%", img)
                                 .Replace("%%WEIGHT%%", weight).Replace("%%HEIGHT%%", height).Replace("%%WIDTH%%", width).Replace("%%LENGTH%%", depth)
                                 .Replace("%%GOOGLECATEGORY%%", googleCategory).Replace("%%PRODUCTTYPE%%", productType).Replace("%%GTIN%%", gtin)
                                 .Replace("%%MPN%%", mpn).Replace("%%PRICE%%", price).Replace("%%EXPDATE%%", expDate).Replace("%%SHIPPINGPRICE%%", fixedShip)
                                 .Replace("%%L0%%", l0).Replace("%%L1%%", l1);
                    }
                }            
            }
            output = output + "\r\n</feed>";
            WriteXML(output);
            file.Close();
            conn.Close();
        }


        private static CsvSubString ReadCsvItem(string line)
        {
            int place;
            string ss;
            int offset = 0;
            
            place = line.IndexOf(",");
            if (line.Substring(0, 1) == "\"")
            {
                place = line.IndexOf("\",");
                offset = 1;
            }
            ss = line.Substring(0, place + offset);
            ss = ss.Replace("\"", "");
            CsvSubString css = new CsvSubString();
            css.Item = ss.Replace("|||***|||", "\"\"");
            css.Remainder = line.Substring(place + 1 + offset, ((line.Length - place) - 1) - offset);
            if (!IsValidXmlString(css.Item)) { css.Item = RemoveInvalidXmlChars(css.Item); }
            return css;
        }

        private static void WriteXML(string text)
        {
            string path = "C:\\feed\\GoogleShopping.xml";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(path))
            {
                file.WriteLine(text);
                file.Close();
            }
            

        }

        private static string FindGoogleCategory(string text, SqlConnection conn)
        {
            string result = "Hardware &gt; Tools";

            SqlCommand command = new SqlCommand("Select top 1 google_product_category gpc from [extracted] where google_category=@gc group by google_product_category, google_category order by google_category, count(id) desc", conn);
            command.Parameters.AddWithValue("@gc", text);
            // int result = command.ExecuteNonQuery();
            using (SqlDataReader reader = command.ExecuteReader())
            {
                if (reader.Read())
                {
                    result = reader["gpc"].ToString();
                }
            }

            return result;
        }

        static string RemoveInvalidXmlChars(string text)
        {
            var validXmlChars = text.Where(ch => XmlConvert.IsXmlChar(ch)).ToArray();
            return new string(validXmlChars);
        }

        static bool IsValidXmlString(string text)
        {
            try
            {
                XmlConvert.VerifyXmlChars(text);
                return true;
            }
            catch
            {
                return false;
            }
        }

            private static DataTable ReadExcel(string filePath, string sheetName)
        {
            DataTable table = new DataTable();
            string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text\"", filePath);
            using (OleDbConnection dbConnection = new OleDbConnection(strConn))
            {
                if (sheetName == null)
                {
                    dbConnection.Open();
                    var dtSchema = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    sheetName = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                }

                using (OleDbDataAdapter dbAdapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", dbConnection)) //rename sheet if required!
                    dbAdapter.Fill(table);
                return table;
            }
        }
    }
    public class CsvSubString
    {
        public CsvSubString()
        {
            Item = null;
            Remainder = null;
        }

        public CsvSubString(string item, string remainder)
            {
                Item = item;
                Remainder = remainder;
            }

            public string Item { get; set; }
            public string Remainder { get; set; }
    }


}
