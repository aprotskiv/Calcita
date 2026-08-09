namespace Calcita.Actions
{
    public sealed class SetRangeDataActionContext
    {
        public SetRangeDataActionContext(object[,] arrayData)
        {
            this.Data = arrayData;
        }

        /// <summary>
        /// data to be set
        /// </summary>
        public object[,] Data { get; }
    }
}