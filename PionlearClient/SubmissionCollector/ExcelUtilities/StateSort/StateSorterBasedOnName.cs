namespace SubmissionCollector.ExcelUtilities.StateSort
{
    internal class StateSorterBasedOnName : BaseStateSorter
    {
        public override void Sort()
        {
            SortColumn = 1;
            base.Sort();
        }
    }
}
