using System;

namespace WopiHost.Abstractions
{
    /// <summary>
    /// Indicate current document is WopiDocs or OtherDocs(email, pdf etc.)
    /// </summary>
    public enum WopiDocsTypeEnum
    {
        WopiDocs,
        OtherDocs
    }
    /// <summary>
    /// Indicate current document is versionable or not
    /// </summary>
    public enum WopiVersionableTypeEnum
    {
        True,
        False
    }
}
