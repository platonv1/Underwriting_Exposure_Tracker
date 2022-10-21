using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace ExposureTracker.Migrations
{
    public partial class lifeUW : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "dbLifeData",
                columns: table => new
                {
                    id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    identifier = table.Column<string>(type: "text", nullable: false),
                    policyno = table.Column<string>(type: "text", nullable: true),
                    firstname = table.Column<string>(type: "text", nullable: true),
                    middlename = table.Column<string>(type: "text", nullable: true),
                    lastname = table.Column<string>(type: "text", nullable: true),
                    fullName = table.Column<string>(type: "text", nullable: true),
                    gender = table.Column<string>(type: "text", nullable: true),
                    clientid = table.Column<string>(type: "text", nullable: true),
                    dateofbirth = table.Column<string>(type: "text", nullable: false),
                    cedingcompany = table.Column<string>(type: "text", nullable: true),
                    cedantcode = table.Column<string>(type: "text", nullable: true),
                    typeofbusiness = table.Column<string>(type: "text", nullable: true),
                    bordereauxfilename = table.Column<string>(type: "text", nullable: true),
                    bordereauxyear = table.Column<int>(type: "integer", nullable: true),
                    soaperiod = table.Column<string>(type: "text", nullable: true),
                    certificate = table.Column<string>(type: "text", nullable: true),
                    plan = table.Column<string>(type: "text", nullable: true),
                    benefittype = table.Column<string>(type: "text", nullable: true),
                    baserider = table.Column<string>(type: "text", nullable: true),
                    currency = table.Column<string>(type: "text", nullable: true),
                    planeffectivedate = table.Column<string>(type: "text", nullable: false),
                    sumassured = table.Column<decimal>(type: "numeric", nullable: false),
                    reinsurednetamountatrisk = table.Column<decimal>(type: "numeric", nullable: false),
                    mortalityrating = table.Column<string>(type: "text", nullable: true),
                    status = table.Column<string>(type: "text", nullable: true),
                    dateuploaded = table.Column<string>(type: "text", nullable: true),
                    uploadedby = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_dbLifeData", x => x.id);
                });

            migrationBuilder.CreateTable(
                name: "dbTranslationTable",
                columns: table => new
                {
                    id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    identifier = table.Column<string>(type: "text", nullable: false),
                    plan_code = table.Column<string>(type: "text", nullable: false),
                    ceding_company = table.Column<string>(type: "text", nullable: true),
                    cedant_code = table.Column<string>(type: "text", nullable: true),
                    benefit_cover = table.Column<string>(type: "text", nullable: true),
                    insured_prod = table.Column<string>(type: "text", nullable: true),
                    prod_description = table.Column<string>(type: "text", nullable: true),
                    base_rider = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_dbTranslationTable", x => x.id);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "dbLifeData");

            migrationBuilder.DropTable(
                name: "dbTranslationTable");
        }
    }
}
