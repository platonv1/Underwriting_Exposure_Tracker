

using ExposureTracker.Models;

namespace ExposureTracker.Data
{
    public class AppDbContext: DbContext
    {

        public AppDbContext (DbContextOptions<AppDbContext> options) : base(options)
        {

        }

        //protected override void OnModelCreating(ModelBuilder builder)
        //{
        //    builder.Entity<TranslationTables>().HasKey(table => new
        //    {
        //        table.id,
        //        table.identifier
        //    });
        //}

        public DbSet<Insured> dbLifeData  { get; set; }
        public DbSet<TranslationTables> dbTranslationTable { get; set; }

        //public DbSet<Policyno> dbpolicyno { get; set; }


    }

 
   



}
