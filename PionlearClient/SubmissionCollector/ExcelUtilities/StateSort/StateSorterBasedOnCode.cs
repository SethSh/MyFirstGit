namespace SubmissionCollector.ExcelUtilities.StateSort
{
    internal class StateSorterBasedOnCode : BaseStateSorter
    {
        public override void Sort()
        {
            SortColumn = 0;
            base.Sort();
        }
    }
}
