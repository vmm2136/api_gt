using AspNet.Security.OAuth.Keycloak;
using Microsoft.IdentityModel.Tokens;
using Minio;
using WebApplication4.Services;


namespace WebApplication4
{
    public class Startup
    {

        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration{ get; }

        public void ConfigureServices(IServiceCollection services)
        {
            services.AddAuthentication(options =>
            {
                options.DefaultAuthenticateScheme = KeycloakAuthenticationDefaults.AuthenticationScheme;
                options.DefaultChallengeScheme = KeycloakAuthenticationDefaults.AuthenticationScheme;
            }).AddJwtBearer(KeycloakAuthenticationDefaults.AuthenticationScheme, options =>
            {
                options.Authority = $"{Configuration.GetValue<String>("Keycloak:BaseUrl")}/realms/plataformagt";
                options.Audience = $"{Configuration.GetValue<String>("Keycloak:Audience")}";
                options.RequireHttpsMetadata = false;

                options.TokenValidationParameters = new Microsoft.IdentityModel.Tokens.TokenValidationParameters
                {
                    ValidAudience = "account"
                };
            });
            services.AddTransient<MinioService>();

            services.AddControllers();

            services.AddHttpContextAccessor();
            services.AddControllers();
            services.AddEndpointsApiExplorer();
            services.AddSwaggerGen();
            services.AddHttpClient();

            services.AddSingleton<MinioService>();
            services.AddScoped<WordService>();
            services.AddScoped<ExcelService>();
            services.AddScoped<PowerPointService>();

        }

        public void Configure(IApplicationBuilder app, IWebHostEnvironment environment)
        {
            if (environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            app.UseHttpsRedirection();

            app.UseRouting();

            app.UseAuthentication();
            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }

    }
}
