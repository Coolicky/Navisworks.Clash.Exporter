using System;
using Autodesk.Navisworks.Api;
using Navisworks.Clash.Exporter.Extensions.Attributes;

namespace Navisworks.Clash.Exporter.Data
{
    [TableName("Comments")]
    public class CommentDto
    {
        public CommentDto(Comment comment, string ownerGuid)
        {
            OwnerGuid = ownerGuid;
            Id = comment.Id;
            Author = comment.Author;
            Body = comment.Body;
            Status = comment.Status.ToString();
            CreationDate = comment.CreationDate;
        }

        [ColumnName("Owner Guid")] public string OwnerGuid { get; }
        [ColumnName("Id")] public long Id { get; }
        [ColumnName("Author")] public string Author { get; }
        [ColumnName("Body")] public string Body { get; }
        [ColumnName("Status")] public string Status { get; }
        [ColumnName("Creation Date")] public DateTime? CreationDate { get; }
    }
}