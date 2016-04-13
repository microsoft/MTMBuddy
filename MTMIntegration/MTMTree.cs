using System.Collections.Generic;

namespace MTMIntegration
{
    public class MTMTree
    {
        public List<MTMTree> Children = new List<MTMTree>();
        private MTMTree parent = new MTMTree();
        private int suiteid;
        private string title = string.Empty;


        public MTMTree()
        {
        }

        public MTMTree(int sid, string name)
        {
            suiteid = sid;
            title = name;
        }

        public MTMTree Addnode(int sid, string name)
        {
            var child = new MTMTree(sid, name);
            child.parent = this;
            Children.Add(child);
            return child;
        }
    }
}