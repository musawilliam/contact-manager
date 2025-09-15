using System;

namespace ContactManager
{
    public class Contact
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public string Surname { get; set; } = "";
        public string Number { get; set; } = "";
        public bool Used { get; set; }
        public DateTime CreatedDate { get; set; } = DateTime.Now;
    }
}