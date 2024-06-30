using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;


namespace WordProcessor.Data
{
    public class NtiDataContext : DbContext
    {
        public NtiDataContext(DbContextOptions<NtiDataContext> dbContextOptions) : base(dbContextOptions) 
        {
            
        }
    }
}
