
#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Windows;
using Application = Autodesk.Revit.ApplicationServices.Application;
using WPF = System.Windows;

#endregion


namespace testPlugin
{
    internal class CustomLevel
    {
        public Level RevitLevel { get; set; }
        public List<CustomSection> SectionsInLevel { get; set; }
        public List<CustomRoom> Rooms { get; set; }
    }
    internal class CustomSection
    {
        public int SectionName { get; set; }
        public CustomLevel LevelOfSection { get; set; }
        public List<CustomApartment> ApartmentsInSection { get; set; }
        public List<CustomRoom> Rooms { get; set; }
    }
    internal class CustomApartment
    {
        public List<CustomRoom> Rooms { get; set; }
        public CustomSection Section { get; set; }
        public int ApartmentNumber { get; set; }
        public int AmountOfRooms { get; set; }

    }
    internal class CustomRoom
    {
        public Room RevitRoom { get; set; }
        public Level Level { get; set; }
        public int Section { get; set; }
        public string Apartment { get; set; }
        public int AmountOfRooms { get; set; }
    }
    //выше объявлены кастомные классы для создания полноценной структуры проекта и удобного доступа
    //с иерархией уровень->все секции на уровне->все квартиры в секции на этаже->все комнаты квартиры
    public sealed partial class ApartmentColoring
    {
        const string AppartmentParam = "ROM_Зона";
        const string SectionParam = "BS_Блок";
        const string AmountOfRoomParam = "ROM_Подзона";
        const string AmountOfRoomIdParam = "ROM_Расчетная_подзона_ID";
        const string AmountOfRoomIndexParam = "ROM_Подзона_Index";
        //объявление констант с строками названиями параметров, чтобы не хардкодить их каждый раз в .LookupParameter
        private bool DoWork(ExternalCommandData commandData, ref String message, ElementSet elements)
        {

            if (null == commandData)
            {

                throw new ArgumentNullException(nameof(commandData));
            }

            if (null == message)
            {

                throw new ArgumentNullException(nameof(message));
            }

            if (null == elements)
            {

                throw new ArgumentNullException(nameof(elements));
            }


            UIApplication ui_app = commandData.Application;
            UIDocument ui_doc = ui_app?.ActiveUIDocument;
            Application app = ui_app?.Application;
            Document doc = ui_doc?.Document;
            try
            {
                using (var tr = new Transaction(doc, "Окрашивание квартир"))//user-friendly название для транзакции 
                {

                    if (TransactionStatus.Started == tr.Start())
                    {

                        var customRoomsCollection = GetAllCustomRooms(doc); // сбор коллекции всех помещений проекта, которые будут анализироваться, и создание для каждого помещения объекта кастомного класса CustomRoom
                        var customLevelCollection = GroupingByLevel(doc,customRoomsCollection); //создание коллекции с группировкой помещений по этажам
                        GroupingBySectionsInLevel(doc, customLevelCollection);//дополнение вышесозданной коллекции и группировка по иерархии уровень-секция-квартира-помещение
                        //после отработки предыдущего метода имеем полностью собранную и подготовленную к закрашиванию коллекцию с выстроенной иерархией
                        var cnt = PaintApartments(customLevelCollection);//метод для покраски стыкующихся помещений 
                        MessageBox.Show("Закрашено следующее количество квартир - " + cnt, "Успешно");
                        return TransactionStatus.Committed == tr.Commit();
                        //в тз это не указано, но я бы уточнил у бимщиков, есть ли гарантии, что необходимые параметры для проверки того, что помещения рядом, будут заполнены
                        //если таких гарантий нет, то я бы предложил дополнительную проверку на геометрическое нахождение стыкующихся квартир, это конечно усложнило бы алгоритм, но добавило бы стабильности плагину
                        //так же предложил бы генерировать, например, xml или excel файл с записями о том, какие помещения были закрашены по той же структуре - этаж-секция-квартира, для отчетности
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Возникла ошибка", ex.Message + "\n" + ex.StackTrace);
            }
            return false;
        }
        
        private List<CustomRoom> GetAllCustomRooms(Document doc)
        {
            List<CustomRoom> customRooms = new List<CustomRoom>();
            var AllRooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).Cast<Room>()
                            .Where(x => x.LookupParameter(AppartmentParam).AsString().ToLower().Contains("квартира")).ToList();
            foreach (var room in AllRooms)
            {
                CustomRoom customRoom = new CustomRoom();
                customRoom.RevitRoom = room;
                customRoom.Level = room.Level;
                customRoom.Section = int.Parse(room.LookupParameter(SectionParam).AsString().Split(' ')[1]); //здесь и далее каст к инту для более удобной работы с этими значениями далее
                //я бы уточнил у проектировщиков есть ли гарантии, что там всегда будет написано значение так, как написано здесь. Если таких гарантий нет то, можно переделать под строки или написать более отказоустойчивый парсер
                customRoom.AmountOfRooms = GetAmountOfRooms(room); //в тз указано: "Сгруппировать помещения по параметру ROM_Подзона, определив количество комнат"
                //сделаю по тз в методе ниже - через параметр ROM_Подзона, но, проанализировав, увидел,спа что в ROM_Расчетная_подзона_ID, по сути хранится та же самая информация
                //и получить размер квартиры через нее удобнее, но я не знаю, какие у вас внутренние порядки и уточнил бы у проектировщика можно ли так)
                //получение количества комнат через него закомментировал
                //customRoom.AmountOfRooms = int.Parse(room.LookupParameter(SectionParam).AsString()[0].ToString());
                customRoom.Apartment = room.LookupParameter(AppartmentParam).AsString().Split(' ')[1]; //здесь то же самое, что и сверху, можно сделать чуть проще через параметр "ROM_Зона_ID"
                //customRoom.Appartment = room.LookupParameter("ROM_Зона_ID").AsString();
                customRooms.Add(customRoom);
            }
            return customRooms;
        }
        
        private int GetAmountOfRooms(Room room)
        {
            var amountParamValue = room.LookupParameter(AmountOfRoomParam).AsString();
            switch (amountParamValue)
            {
                case "Однокомнатная квартира":
                    return 1;
                case "Двухкомнатная квартира":
                    return 2;
                case "Трехкомнатная квартира":
                    return 3;
                case "Четырехкомнатная квартира":
                    return 4;
                case "Пятикомнатная квартира":
                    return 5;
                default:
                    return 0;
                    //сделана конструкция switch-case на случай, если в проекте будут квартиры с большим количеством помещений и было бы удобно их сюда добавить
                    
            }
        }
        private List<CustomLevel> GroupingByLevel(Document doc, List<CustomRoom> customRooms)
        {
            List<CustomLevel> customLevels = new List<CustomLevel>();
            var levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            foreach(var level in levels)
            {
                CustomLevel customLevel = new CustomLevel();
                List<CustomRoom> roomsOnLevel = customRooms.Where(x=>x.Level.Id==level.Id).ToList();
                customLevel.RevitLevel = level;
                customLevel.Rooms = roomsOnLevel;
                customLevels.Add(customLevel);
            }
            return customLevels;

        }

        private void GroupingBySectionsInLevel(Document doc, List<CustomLevel> customLevels)
        {
            foreach(var level in customLevels)
            {
                var sectionsOnLevel = level.Rooms.Select(x=>x.Section).ToList().Distinct(); //нахождение секций, которые есть на этаже
                List<CustomSection> customSections = new List<CustomSection>();
                foreach(var section in sectionsOnLevel)
                {
                    CustomSection customSection = new CustomSection();
                    customSection.LevelOfSection = level;
                    customSection.SectionName = section;
                    customSection.Rooms = level.Rooms.Where(x => x.Section == section).ToList();
                    customSection.ApartmentsInSection = GroupingByApartmentsInSection(doc, customSection); //группировка комнат на секции по квартирам
                    customSections.Add(customSection);

                    
                }
                level.SectionsInLevel = customSections;
            }
        }
        private List<CustomApartment> GroupingByApartmentsInSection(Document doc, CustomSection customSection)
        {
            var apartmentsInSection = customSection.Rooms.Select(x => x.Apartment).ToList().Distinct();
            List<CustomApartment> customApartments = new List<CustomApartment>();
            foreach (var ap in apartmentsInSection)
            {
                CustomApartment customApartment = new CustomApartment();
                customApartment.Section = customSection;
                customApartment.ApartmentNumber = int.Parse(ap);
                customApartment.Rooms = customSection.Rooms.Where(x=>x.Apartment == ap).ToList();
                customApartment.AmountOfRooms = customApartment.Rooms.FirstOrDefault().AmountOfRooms;
                customApartments.Add(customApartment);
            }
            return customApartments;
            
        }
        private int PaintApartments(List<CustomLevel> customLevels)
        {
            int cnt = 0; //счетчик для вывода информации пользователю о количестве закрашенных квартир
            foreach (var level in customLevels)
            {
                foreach (var section in level.SectionsInLevel)
                {
                    bool prevWasPainted = false; //булевская переменная для обработки случаев, когда рядом находятся например 3 и более квартир, чтобы они все не перекрашивались
                    //в проекте есть такой пример на первой секции второго этажа
                    //то есть реализована закраска через одну квартиру в таких случаях
                    foreach (var ap in section.ApartmentsInSection)
                    {
                        int currentAp = ap.ApartmentNumber;
                        int amountOfRoomsInCurrentAp = ap.AmountOfRooms;
                        CustomApartment nextAp = section.ApartmentsInSection.Where(x => x.ApartmentNumber == currentAp + 1).FirstOrDefault();
                        if (nextAp == null)
                            continue;//для обработки последней квартиры на этаже
                        if (nextAp.AmountOfRooms == amountOfRoomsInCurrentAp)
                        {
                            if (prevWasPainted)
                            {
                                foreach (var room in nextAp.Rooms)
                                {
                                    var AmountOfRoomIdParamValue = room.RevitRoom.LookupParameter(AmountOfRoomIdParam).AsString();
                                    room.RevitRoom.LookupParameter(AmountOfRoomIndexParam).Set(AmountOfRoomIdParamValue + ".Полутон");
                                }
                                cnt++;
                                prevWasPainted = true;
                            }
                            else
                            {
                                foreach (var room in ap.Rooms)
                                {
                                    var AmountOfRoomIdParamValue = room.RevitRoom.LookupParameter(AmountOfRoomIdParam).AsString();
                                    room.RevitRoom.LookupParameter(AmountOfRoomIndexParam).Set(AmountOfRoomIdParamValue + ".Полутон");
                                }
                                cnt++;
                                prevWasPainted = true;
                            }
                        }
                        else
                        {
                            prevWasPainted = false;
                        }
                    }
                }
            }
            return cnt;
        }
    }
}

