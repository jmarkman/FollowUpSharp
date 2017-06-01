namespace FollowUpSharp
{
    public class IMSEntry
    {
        public string ControlNum { get; } = "";
        public string DueDate { get; } = "";
        public string CreatedDate { get; } = "";
        public string EffectiveDate { get; } = "";
        public string Broker { get; } = "";
        public string BrokerCompany { get; } = "";
        public string FirstName { get; } = "";
        public string LastName { get; } = "";
        public string BrokerEmail { get; set; } = "";
        public string UnderwriterFirstName { get; } = "";
        public string UnderwriterLastName { get; } = "";
        public string UnderwriterEmail { get; } = "";
        public string NamedInsured { get; } = "";

        /// <summary>
        /// Constructor for an IMSEntry object. This object contains an abbreviated insured row in
        /// the SQL database, which contains pertinent information for present & future parties.
        /// This is an object that will represent the equivalent row in SQL which will
        /// be written to the record-keeping Excel document.
        /// </summary>
        /// <param name="controlNum">The control number associated with the IMSEntry's account</param>
        /// <param name="dueDate"></param>
        /// <param name="createdDate"></param>
        /// <param name="effectiveDate"></param>
        /// <param name="broker"></param>
        /// <param name="brokerCompany"></param>
        /// <param name="firstName">The first name of the broker associated with the IMSEntry</param>
        /// <param name="lastName"></param>
        /// <param name="brokerEmail">The email of the broker associated with the IMSEntry</param>
        /// <param name="underwriterFirstName"></param>
        /// <param name="underwriterLastName"></param>
        /// <param name="underwriterEmail">The email of the underwriter associated with the IMSEntry</param>
        /// <param name="namedInsured">The name of the IMSEntry's account as it is in IMS</param>     
        public IMSEntry(string controlNum = "", string dueDate = "", string effectiveDate = "", string createdDate = "", string namedInsured = "", string broker = "", string brokerCompany = "", string firstName = "", string lastName = "", string brokerEmail = "", string underwriterFirstName = "", string underwriterLastName = "", string underwriterEmail = "")
        {
            ControlNum = controlNum;
            DueDate = dueDate;
            EffectiveDate = effectiveDate;
            CreatedDate = createdDate;
            NamedInsured = namedInsured;
            Broker = broker;
            BrokerCompany = brokerCompany;
            FirstName = firstName;
            LastName = lastName;
            BrokerEmail = brokerEmail;
            UnderwriterFirstName = underwriterFirstName;
            UnderwriterLastName = underwriterLastName;
            UnderwriterEmail = underwriterEmail;
        }
    }
}