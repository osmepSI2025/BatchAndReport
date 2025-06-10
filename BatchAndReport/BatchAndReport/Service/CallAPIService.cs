using BatchAndReport.Models;

using BatchAndReport.Repository;
using System;
using System.IdentityModel.Tokens.Jwt;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace BatchAndReport.Services
{
    public class CallAPIService : ICallAPIService
    {
        private readonly HttpClient _httpClient;
        private readonly string Api_ErrorLog;
        private readonly string Api_SysCode;

        private readonly IApiInformationRepository _repositoryApi;
        public CallAPIService(HttpClient httpClient, IConfiguration configuration, IApiInformationRepository repositoryApi)
        {
            _httpClient = httpClient;
            Api_ErrorLog = configuration["Information:Api_ErrorLog"] ?? throw new ArgumentNullException("Api_ErrorLog is missing in appsettings.json");
            Api_SysCode = configuration["Information:Api_SysCode"] ?? throw new ArgumentNullException("Api_SysCode is missing in appsettings.json");

            _repositoryApi = repositoryApi ?? throw new ArgumentNullException(nameof(repositoryApi));
        }

        public async Task<string> GetDataApiAsync_Login(MapiInformationModels apiModels)
        {
            if (string.IsNullOrEmpty(apiModels.Urlproduction))
                throw new ArgumentException("Urlproduction cannot be null or empty.");
            if (string.IsNullOrEmpty(apiModels.Username) || string.IsNullOrEmpty(apiModels.Password))
                throw new ArgumentException("Username and Password cannot be null or empty.");

            string requestJson = "";
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                WriteIndented = true
            };

            try
            {
                using var request = new HttpRequestMessage(HttpMethod.Post, apiModels.Urlproduction);

                // Prepare the payload for the token request
                var payload = new
                {
                    username = apiModels.Username,
                    password = apiModels.Password
                };

                // Serialize the payload to JSON and set it as the request content
                requestJson = JsonSerializer.Serialize(payload, options);
                request.Content = new StringContent(requestJson, Encoding.UTF8, "application/json");

                // Send the request using the existing _httpClient
                var response = await _httpClient.SendAsync(request);

                // Ensure the response is successful
                response.EnsureSuccessStatusCode();

                // Read the response content
                var content = await response.Content.ReadAsStringAsync();

                // Parse the JSON response to extract the token
                using var doc = JsonDocument.Parse(content);
                if (!doc.RootElement.TryGetProperty("access_token", out var tokenElement))
                    throw new Exception("Token not found in response");

                // update token
                var token = tokenElement.GetString();
                _repositoryApi.UpdateAllBearerTokensAsync(token);

                return tokenElement.GetString() ?? throw new Exception("Token is null or empty");
            }
            catch (Exception ex)
            {
                //var errorLog = new ErrorLogModels
                //{
                //    Message = "Function " + apiModels.ServiceNameTh + " " + ex.Message,
                //    StackTrace = ex.StackTrace,
                //    Source = ex.Source,
                //    TargetSite = ex.TargetSite?.ToString(),
                //    ErrorDate = DateTime.Now,
                //    UserName = apiModels.Username, // ดึงจาก context หรือ session
                //    Path = apiModels.Urlproduction,
                //    HttpMethod = apiModels.MethodType,
                //    RequestData = requestJson, // serialize เป็น JSON
                //    InnerException = ex.InnerException?.ToString(),
                //    SystemCode = Api_SysCode,
                //    CreatedBy = "system"

                //};
                //await RecErrorLogApiAsync(apiModels, errorLog);
                throw new Exception("Error in GetData: " + ex.Message + " | Inner Exception: " + ex.InnerException?.Message);
            }
        }

        public async Task RecErrorLogApiAsync(MapiInformationModels apiModels, ErrorLogModels eModels)
        {

            // ✅ เลือก URL ตาม _FlagDev
            string? apiUrl = Api_ErrorLog;



            try
            {
                using var request = new HttpRequestMessage(HttpMethod.Post, apiUrl);

                // ✅ ใส่ API Key ถ้ามี
                if (!string.IsNullOrEmpty(apiModels.ApiKey))
                    request.Headers.Add("X-Api-Key", apiModels.ApiKey);

                // ✅ ใส่ Basic Authentication ถ้ามี Username & Password
                if (!string.IsNullOrEmpty(apiModels.Username) && !string.IsNullOrEmpty(apiModels.Password))
                {
                    var authHeader = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{apiModels.Username}:{apiModels.Password}"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Basic", authHeader);
                }

                // ✅ แปลง SendData เป็น JSON และแนบไปกับ Body ของ Request
                var jsonData = JsonSerializer.Serialize(eModels);
                request.Content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                // ✅ Call API และรอผลลัพธ์
                using var response = await _httpClient.SendAsync(request);
                string responseData = await response.Content.ReadAsStringAsync();

            }
            catch (Exception ex)
            {

            }
        }
        public async Task<T?> ReadJsonFileAsync<T>(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"The file '{filePath}' was not found.");
            }

            var jsonContent = await File.ReadAllTextAsync(filePath);
            return JsonSerializer.Deserialize<T>(jsonContent);
        }
        public bool IsTokenExpired(string token)
        {
            var jwtToken = new JwtSecurityTokenHandler().ReadToken(token) as JwtSecurityToken;
            if (jwtToken == null) return true;

            var expiry = jwtToken.ValidTo;
            return expiry < DateTime.UtcNow;
        }
        public async Task<string> GetTokenAsync(string username, string password, string tokenUrl)
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, tokenUrl);

            // Prepare the payload for the token request
            var payload = new
            {
                username = username,
                password = password
            };

            // Serialize the payload to JSON and set it as the request content
            request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

            try
            {
                // Send the request using the existing _httpClient
                var response = await _httpClient.SendAsync(request);

                // Ensure the response is successful
                response.EnsureSuccessStatusCode();

                // Read the response content
                var content = await response.Content.ReadAsStringAsync();

                // Parse the JSON response to extract the token
                using var doc = JsonDocument.Parse(content);
                var token = doc.RootElement.GetProperty("access_token").GetString();

                return token ?? throw new Exception("Token not found in response");
            }
            catch (Exception ex)
            {
                // Log the error or handle it as needed
                throw new Exception($"Failed to retrieve token: {ex.Message}", ex);
            }
        }
        public async Task<string> GetDataApiAsync(MapiInformationModels apiModels, object xdata)
        {
            string requestJson = "";
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                WriteIndented = true
            };

            try
            {
                using var handler = new HttpClientHandler
                {
                    ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
                };

                using var httpClient = new HttpClient(handler);

                // Construct URL
                var url = apiModels.Urlproduction ?? throw new Exception("URL is missing.");
                requestJson = url;

                // Create request
                var request = new HttpRequestMessage(
                    apiModels.MethodType == "POST" ? HttpMethod.Post : HttpMethod.Get,
                    url
                );

                // Set Authorization Headers
                if (apiModels.AuthorizationType == "Basic")
                {
                    var byteArray = System.Text.Encoding.ASCII.GetBytes($"{apiModels.Username}:{apiModels.Password}");
                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                }
                else if (apiModels.AuthorizationType == "Bearer")
                {
                    string? token = apiModels.Bearer;

                    // Check if token is expired
                    if (!string.IsNullOrEmpty(token) && IsTokenExpired(token))
                    {
                        var LApi = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "user-api" });

                        var apiParamx = LApi.Select(x => new MapiInformationModels
                        {
                            ServiceNameCode = x.ServiceNameCode,
                            ApiKey = x.ApiKey,
                            AuthorizationType = x.AuthorizationType,
                            ContentType = x.ContentType,
                            CreateDate = x.CreateDate,
                            Id = x.Id,
                            MethodType = x.MethodType,
                            ServiceNameTh = x.ServiceNameTh,
                            Urldevelopment = x.Urldevelopment,
                            Urlproduction = x.Urlproduction,
                            Username = x.Username,
                            Password = x.Password,
                            UpdateDate = x.UpdateDate,
                            Bearer = x.Bearer,
                        }).FirstOrDefault(); // Use FirstOrDefault to handle empty lists


                        var result = await GetDataApiAsync_Login(apiParamx);
                        token = result;
                    }

                    if (string.IsNullOrEmpty(token))
                    {
                        throw new Exception("Bearer token is missing or expired.");
                    }

                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                }
                else if (apiModels.AuthorizationType == "ApiKey")
                {
                    request.Headers.Add("X-Api-Key", apiModels.ApiKey);
                }
                else
                {
                    throw new Exception("Authorization type not supported");
                }

                // Add JSON body for POST requests
                if (apiModels.MethodType == "POST")
                {
                    if (xdata != null)
                    {
                        var jsonBody = JsonSerializer.Serialize(xdata, options);
                        requestJson = jsonBody; // For logging
                        request.Content = new StringContent(jsonBody, System.Text.Encoding.UTF8, "application/json");
                    }
                    else
                    {
                        throw new Exception("POST request must have a JSON body.");
                    }
                }

                // Send Request
                var response = await httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();
                var content = await response.Content.ReadAsStringAsync();
                var jsonNode = JsonNode.Parse(content);

                
                return jsonNode?.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) ?? "{}";
            }
            catch (Exception ex)
            {
                var errorLog = new ErrorLogModels
                {
                    Message = "Function " + apiModels.ServiceNameTh + " " + ex.Message,
                    StackTrace = ex.StackTrace,
                    Source = ex.Source,
                    TargetSite = ex.TargetSite?.ToString(),
                    ErrorDate = DateTime.Now,
                    UserName = apiModels.Username, // ดึงจาก context หรือ session
                    Path = apiModels.Urlproduction,
                    HttpMethod = apiModels.MethodType,
                    RequestData = requestJson, // serialize เป็น JSON
                    InnerException = ex.InnerException?.ToString(),
                    SystemCode = Api_SysCode,
                    CreatedBy = "system"

                };
                await RecErrorLogApiAsync(apiModels, errorLog);
                throw new Exception("Error in GetData: " + ex.Message + " | Inner Exception: " + ex.InnerException?.Message);
            }
        }

        public async Task<string> GetDataByParamApiAsync(MapiInformationModels apiModels, string typeValue)
        {
            string requestJson = "";
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                WriteIndented = true
            };

            try
            {
                using var handler = new HttpClientHandler
                {
                    ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
                };

                using var httpClient = new HttpClient(handler);

                // Construct URL
                var url = apiModels.Urlproduction ?? throw new Exception("URL is missing.");
                if (url.Contains("{EmpId}"))
                {
                    url = url.Replace("{EmpId}", typeValue);
                }

                if (url.Contains("{BudYear}"))
                {
                    url = url.Replace("{BudYear}", typeValue);
                }

                requestJson = url;

                // Create request
                var request = new HttpRequestMessage(
                    apiModels.MethodType == "POST" ? HttpMethod.Post : HttpMethod.Get,
                    requestJson
                );

                // Set Authorization Headers
                if (apiModels.AuthorizationType == "Basic")
                {
                    var byteArray = System.Text.Encoding.ASCII.GetBytes($"{apiModels.Username}:{apiModels.Password}");
                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                }
                else if (apiModels.AuthorizationType == "Bearer")
                {
                    string? token = apiModels.Bearer;

                    // Check if token is expired
                    if (!string.IsNullOrEmpty(token) && IsTokenExpired(token))
                    {
                        var LApi = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "user-api" });

                        var apiParamx = LApi.Select(x => new MapiInformationModels
                        {
                            ServiceNameCode = x.ServiceNameCode,
                            ApiKey = x.ApiKey,
                            AuthorizationType = x.AuthorizationType,
                            ContentType = x.ContentType,
                            CreateDate = x.CreateDate,
                            Id = x.Id,
                            MethodType = x.MethodType,
                            ServiceNameTh = x.ServiceNameTh,
                            Urldevelopment = x.Urldevelopment,
                            Urlproduction = x.Urlproduction,
                            Username = x.Username,
                            Password = x.Password,
                            UpdateDate = x.UpdateDate,
                            Bearer = x.Bearer,
                        }).FirstOrDefault(); // Use FirstOrDefault to handle empty lists


                        var result = await GetDataApiAsync_Login(apiParamx);
                        token = result;
                    }

                    if (string.IsNullOrEmpty(token))
                    {
                        throw new Exception("Bearer token is missing or expired.");
                    }

                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                }
                else if (apiModels.AuthorizationType == "ApiKey")
                {
                    request.Headers.Add("X-Api-Key", apiModels.ApiKey);
                }
                else
                {
                    throw new Exception("Authorization type not supported");
                }

                // Send Request
                var response = await httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();
                var content = await response.Content.ReadAsStringAsync();
                var jsonNode = JsonNode.Parse(content);


                return jsonNode?.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) ?? "{}";
            }
            catch (Exception ex)
            {
                var errorLog = new ErrorLogModels
                {
                    Message = "Function " + apiModels.ServiceNameTh + " " + ex.Message,
                    StackTrace = ex.StackTrace,
                    Source = ex.Source,
                    TargetSite = ex.TargetSite?.ToString(),
                    ErrorDate = DateTime.Now,
                    UserName = apiModels.Username, // ดึงจาก context หรือ session
                    Path = apiModels.Urlproduction,
                    HttpMethod = apiModels.MethodType,
                    RequestData = requestJson, // serialize เป็น JSON
                    InnerException = ex.InnerException?.ToString(),
                    SystemCode = Api_SysCode,
                    CreatedBy = "system"

                };
                await RecErrorLogApiAsync(apiModels, errorLog);
                throw new Exception("Error in GetData: " + ex.Message + " | Inner Exception: " + ex.InnerException?.Message);
            }
        }

        public async Task<string> GetDataEmpMovementApiAsync(MapiInformationModels apiModels, string EmpId)
        {
            string requestJson = "";
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                WriteIndented = true
            };

            try
            {
                using var handler = new HttpClientHandler
                {
                    ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
                };

                using var httpClient = new HttpClient(handler);

                // Construct URL
                var url = apiModels.Urlproduction ?? throw new Exception("URL is missing.");
                url = url.Replace("{EmpId}", EmpId);
                

                requestJson = url;

                // Create request
                var request = new HttpRequestMessage(
                    apiModels.MethodType == "POST" ? HttpMethod.Post : HttpMethod.Get,
                    requestJson
                );

                // Set Authorization Headers
                if (apiModels.AuthorizationType == "Basic")
                {
                    var byteArray = System.Text.Encoding.ASCII.GetBytes($"{apiModels.Username}:{apiModels.Password}");
                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                }
                else if (apiModels.AuthorizationType == "Bearer")
                {
                    string? token = apiModels.Bearer;

                    // Check if token is expired
                    if (!string.IsNullOrEmpty(token) && IsTokenExpired(token))
                    {
                        var LApi = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "user-api" });

                        var apiParamx = LApi.Select(x => new MapiInformationModels
                        {
                            ServiceNameCode = x.ServiceNameCode,
                            ApiKey = x.ApiKey,
                            AuthorizationType = x.AuthorizationType,
                            ContentType = x.ContentType,
                            CreateDate = x.CreateDate,
                            Id = x.Id,
                            MethodType = x.MethodType,
                            ServiceNameTh = x.ServiceNameTh,
                            Urldevelopment = x.Urldevelopment,
                            Urlproduction = x.Urlproduction,
                            Username = x.Username,
                            Password = x.Password,
                            UpdateDate = x.UpdateDate,
                            Bearer = x.Bearer,
                        }).FirstOrDefault(); // Use FirstOrDefault to handle empty lists


                        var result = await GetDataApiAsync_Login(apiParamx);
                        token = result;
                    }

                    if (string.IsNullOrEmpty(token))
                    {
                        throw new Exception("Bearer token is missing or expired.");
                    }

                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                }
                else if (apiModels.AuthorizationType == "ApiKey")
                {
                    request.Headers.Add("X-Api-Key", apiModels.ApiKey);
                }
                else
                {
                    throw new Exception("Authorization type not supported");
                }

                // Send Request
                var response = await httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();
                var content = await response.Content.ReadAsStringAsync();
                var jsonNode = JsonNode.Parse(content);


                return jsonNode?.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) ?? "{}";
            }
            catch (Exception ex)
            {
                var errorLog = new ErrorLogModels
                {
                    Message = "Function " + apiModels.ServiceNameTh + " " + ex.Message,
                    StackTrace = ex.StackTrace,
                    Source = ex.Source,
                    TargetSite = ex.TargetSite?.ToString(),
                    ErrorDate = DateTime.Now,
                    UserName = apiModels.Username, // ดึงจาก context หรือ session
                    Path = apiModels.Urlproduction,
                    HttpMethod = apiModels.MethodType,
                    RequestData = requestJson, // serialize เป็น JSON
                    InnerException = ex.InnerException?.ToString(),
                    SystemCode = Api_SysCode,
                    CreatedBy = "system"

                };
                await RecErrorLogApiAsync(apiModels, errorLog);
                throw new Exception("Error in GetData: " + ex.Message + " | Inner Exception: " + ex.InnerException?.Message);
            }
        }

    }
}
