using System.Collections;

namespace GBTH.List
{
    public class ListViewColumnSorter : IComparer
    {
        private int column_to_sort;
        private SortOrder order_of_sort;
        private CaseInsensitiveComparer object_compare;
        private Type type;

        public int SortColumn
        {
            set
            {
                this.column_to_sort = value;
            }
            get
            {
                return column_to_sort;
            }
        }
        public SortOrder Order
        {
            set
            {
                this.order_of_sort = value;
            }
            get
            {
                return order_of_sort;
            }
        }

        public ListViewColumnSorter()
        {
            this.column_to_sort = 0;
            this.order_of_sort = SortOrder.None;
            this.object_compare = new CaseInsensitiveComparer();
            this.type = typeof(string);
        }

        public int Compare(object? x, object? y)
        {
            int result;

            if (this.type == typeof(int))
            {
                result = this.object_compare.Compare(int.Parse((x as ListViewItem)!.SubItems[column_to_sort].Text),
                    int.Parse((y as ListViewItem)!.SubItems[column_to_sort].Text));
            }
            else
            {
                result = this.object_compare.Compare((x as ListViewItem)!.SubItems[column_to_sort].Text,
                    (y as ListViewItem)!.SubItems[column_to_sort].Text);
            }

            if (this.order_of_sort == SortOrder.Ascending)
            {
                return result;
            }
            else if (this.order_of_sort == SortOrder.Descending)
            {
                return (-result);
            }
            else
            {
                return 0;
            }
        }

        public void SetType(Type type)
        {
            this.type = type;
        }
    }
}
