namespace ReportsLibraryOpenXmlSdk.Entities
{
    /// <summary>
    /// Класс создания объекта строки подключения к БД MSSQL
    /// </summary>
    /// <param name="server">Название сервера</param>
    /// <param name="nameDb">Название базы данных на сервере</param>
    /// <param name="user">Логин</param>
    /// <param name="password">Пароль</param>
    public class ConnectionConfigMssql(string server, string nameDb, string user, string password)
    {
        /// <summary>
        /// Название сервера
        /// </summary>
        public string Server { get; set; } = server;

        /// <summary>
        /// Название базы данных на сервере
        /// </summary>
        public string NameDb { get; set; } = nameDb;

        /// <summary>
        /// Логин
        /// </summary>
        public string User { get; set; } = user;

        /// <summary>
        /// Пароль
        /// </summary>
        public string Password { get; set; } = password;
    }
}
