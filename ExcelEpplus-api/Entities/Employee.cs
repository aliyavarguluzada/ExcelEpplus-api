namespace ExcelEpplus_api.Entities
{
    public class Employee
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Age { get; set; }
        public bool isActive { get; set; }
        public bool isDeleted { get; set; }
        public int ReqUserId { get; set; }
        public DateTime ReqDate { get; set; }
        public DateTime UpdateDate { get; set; }
    }
}
