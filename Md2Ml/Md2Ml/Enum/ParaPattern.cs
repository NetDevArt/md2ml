namespace Md2Ml.Enum
{
    /// <summary>
    /// This enumerator makes it possible to reference the different types of "paragraphs" existing in markdown.
    /// </summary>
    public enum ParaPattern
    {
        Image = 1,
        Table = 2,
        TableHeaderSeparation = 3,
        Quote = 4,
        InfiniteHeading = 5,
        UnorderedList = 6,
        OrderedList = 7,
        CodeBlock = 8,
        ReqTitle = 9,
        ReqProperties1 = 10,
        ReqProperties2 = 11,
        AnyChar = 12
    }
}
