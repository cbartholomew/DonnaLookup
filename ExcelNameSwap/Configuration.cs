using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNameSwap
{
    public class Configuration
    {
        public const string file_path = @"C:\Users\Christopher\Source\Repos\DonnaLookup\ExcelNameSwap\CSE Fall All Call 2016v1.xlsx";

        public enum COLUMNS
        {
            ID = 0,
            Household_ID = 1,
            Entity_type = 2,
            Status = 3,
            Sort_name = 4,
            first_name = 5,
            middle_name = 6,
            last_name = 7,
            pref_mailing_name = 8,
            job_title = 9,
            employer_name = 10,
            spouse_name = 11,
            spouse_entity_type = 12,
            spouse_id_number = 13,
            pref_city = 14,
            pref_State = 15,
            email_contact = 16,
            receive_nothing_from1 = 17,
            email_address = 18,
            spouse_email_address = 19,
            email_address1 = 20,
            email_address2 = 21,
            email_address3 = 22,
            email_address4 = 23,
            email_address5 = 24,
            email_address6 = 25
        }

        public enum EMAIL_WEIGHT
        { 
            UW              = 0,
            CS              = 0,
            MATH            = 0,
            AMAZON          = 0,
            OTHER           = 1,
            HOTMAIL         = 2,
            YAHOO           = 3,
            LIVE            = 4,
            GMAIL           = 5,
        }

        public static EMAIL_WEIGHT GetWeight(string email)
        {
            if (email == "UW" || email == "CS" || email == "MATH" || email == "AMAZON")
            {
                return EMAIL_WEIGHT.UW;
            }
            switch (email)
            {
                case "GMAIL":
                    return EMAIL_WEIGHT.GMAIL;                 
                case "LIVE":
                    return EMAIL_WEIGHT.LIVE;                  
                case "YAHOO":
                    return EMAIL_WEIGHT.YAHOO;
                case "HOTMAIL":
                    return EMAIL_WEIGHT.HOTMAIL;                 
                default:
                    return EMAIL_WEIGHT.OTHER;
            }
        
        }
    }
}
