var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Register the ObakService
builder.Services.AddSingleton<IObakService, ObakService>(provider =>
    new ObakService("mongodb+srv://anubis1080p:vHNelDfzCdi0eR3W@obak.iyvad.mongodb.net/?retryWrites=true&w=majority&appName=OBAK"));

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
