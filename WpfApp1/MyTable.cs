using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1
{
    public class MyTable
    {
        public MyTable()
        {

        }
        public MyTable(string Id, string Name, string Description, string Source, string ObjectVozd, string Conf, string Celostn, string Access, string Date, string LastChange)
        {
            this.Id = Id;
            this.Name = Name;
            this.Description = Description;
            this.Source = Source;
            this.ObjectVozd = ObjectVozd;
            this.Conf = Conf;
            this.Celostn = Celostn;
            this.Access = Access;
            this.Date = Date;
            this.LastChange = LastChange;
        }

        public override string ToString()
        {
            if (Conf == "1")
                Conf = "Да";
            else if (Conf == "0")
                Conf = "Нет";
            if (Access == "1")
                Access = "Да";
            else if (Access == "0")
                Access = "Нет";
            if (Celostn == "1")
                Celostn = "Да";
            else if (Celostn == "0")
                Celostn = "Нет";
            return $"Идентификатор УБИ = {Id} \n\rНаименование УБИ = {Name}\n\rОписание = {Description}\n\rИсточник угрозы = {Source}\n\rОбъект воздействия = {ObjectVozd}\n\rНарушение конфиденциальности = {Conf}\n\rНарушение целостности = {Celostn}\n\rНарушение доступности = {Access}\n\rДата включения угрозы в БнД = {Date}\n\rДата последнего изменения = {LastChange}";
        }

        public MyTable(string Ubi, string Name)
        {
            this.Ubi = Ubi;
            this.Name = Name;
        }
        public string Ubi { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Source { get; set; }
        public string ObjectVozd { get; set; }
        public string Conf { get; set; }
        public string Celostn { get; set; }
        public string Access { get; set; }
        public string Date { get; set; }
        public string LastChange { get; set; }
    }

    
}
