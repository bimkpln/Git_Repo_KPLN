using System.ComponentModel.DataAnnotations;

namespace KPLN_TaskManager.Common
{
    /// <summary>
    /// Контейнер для изображения элемента (массив байтов (BLOB)).
    /// Создаётся в отдельные БД, чтобы добавлять несколько картинок без нагрузки на БД (при большом количестве - торомзит)
    /// </summary>
    public sealed class TaskEntity_ImageBuffer
    {
        public TaskEntity_ImageBuffer()
        {
        }

        public TaskEntity_ImageBuffer(int id, int taskEntityId, byte[] imageBuffer)
        {
            Id = id;
            TaskEntityId = taskEntityId;
            ImageBuffer = imageBuffer;
        }

        #region Данные из БД
        [Key]
        public int Id { get; set; }
        
        /// <summary>
        /// Ссылка на ID из БД по TaskEntity
        /// </summary>
        public int TaskEntityId { get; set; }

        /// <summary>
        /// Изображение элемента (массив байтов (BLOB))
        /// </summary>
        public byte[] ImageBuffer { get; set; }
        #endregion
    }
}
