namespace Md2Ml.Enum
{
    /// <summary>
    /// This enumerator makes it possible to reference the different types of "paragraphs" existing in markdown.
    /// </summary>
    public enum ParaPattern
    {
        Image,
        Table,
        TableHeaderSeparation,
        Quote,
        InfiniteHeading,
        UnorderedList,
        OrderedList,
        CodeBlock,
        ReqTitle,
        ReqProperties1,
        ReqProperties2,
        AnyChar
    }
}
