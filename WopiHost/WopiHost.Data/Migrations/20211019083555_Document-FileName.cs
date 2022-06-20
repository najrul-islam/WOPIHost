using Microsoft.EntityFrameworkCore.Migrations;

namespace WopiHost.Data.Migrations
{
    public partial class DocumentFileName : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "FileName",
                table: "Document",
                type: "NVARCHAR(250)",
                maxLength: 250,
                nullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "FileName",
                table: "Document");
        }
    }
}
