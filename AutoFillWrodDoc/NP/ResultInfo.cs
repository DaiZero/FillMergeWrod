using System.Collections.Generic;

namespace AutoFillWrodDoc
{
    public class ResultInfo
    {
        public bool Succeed { get; set; }

        public string Message { get; set; }

        public WordDataTemplate WordDataTemplate { get; set; }

        public List<ChannelWordInfo> ChannelWordInfos { get; set; }
    }
}
