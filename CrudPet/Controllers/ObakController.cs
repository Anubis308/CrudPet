using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Threading.Tasks;

namespace CrudPet.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ObakController : ControllerBase
    {
        private readonly ILogger<ObakController> _logger;
        private readonly IObakService _obakService;

        public ObakController(ILogger<ObakController> logger, IObakService obakService)
        {
            _logger = logger;
            _obakService = obakService;
        }

        [HttpGet("GetData")]
        public async Task<IActionResult> GetData()
        {
            const string filePath = @"H:\Projects\OBAK\Obak.xlsx";
            try
            {
                var result = await _obakService.GetAndInsertDataAsync(filePath);
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occurred while getting data.");
                return BadRequest(ex.Message);
            }
        }
    }
}
