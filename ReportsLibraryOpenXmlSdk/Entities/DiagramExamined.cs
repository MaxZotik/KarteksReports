namespace ReportsLibraryOpenXmlSdk.Entities
{
    /// <summary>
    /// Класс схемы подключния оборудвания в режиме "Подъем"
    /// </summary>
    /// <param name="gearboxID">ID оборудования из таблицы "Gearboxes" в БД</param>
    /// <param name="point">Номер точки подключния</param>
    /// <param name="axisXmvk">Номер МВК оси Х (вертикальное)</param>
    /// <param name="axisXchannel">Номер канала оси Х (вертикальное)</param>
    /// <param name="axisYmvk">Номер МВК оси Y (горизонтальное)</param>
    /// <param name="axisYchannel">Номер канала оси Y (горизонтальное)</param>
    /// <param name="axisZmvk">Номер МВК оси Z (осевое)</param>
    /// <param name="axisZchannel">Номер канала оси Z (осевое)</param>
    public class DiagramExamined(int gearboxID, int point, int axisXmvk, int axisXchannel, int axisYmvk, int axisYchannel, int axisZmvk, int axisZchannel) : Diagram(point, axisXmvk, axisXchannel, axisYmvk, axisYchannel, axisZmvk, axisZchannel)
    {
        /// <summary>
        /// ID оборудования из таблицы "Gearboxes" в БД
        /// </summary>
        public int GearboxID { get; set; } = gearboxID;
    }
}
