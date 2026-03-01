using System;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Azure.Identity;

namespace MSBookingsApp
{
    /// <summary>
    /// Sample application to interact with Microsoft Bookings via Microsoft Graph API
    /// </summary>
    class Program
    {
        // Your Azure AD App Registration details
        private const string TenantId = "YOUR_TENANT_ID";
        private const string ClientId = "YOUR_CLIENT_ID";
        private const string ClientSecret = "YOUR_CLIENT_SECRET"; // Use certificate in production!

        static async Task Main(string[] args)
        {
            Console.WriteLine("=== Microsoft Bookings API Sample ===\n");

            try
            {
                // Initialize Graph client
                var graphClient = CreateGraphClient();

                // Test 1: List all booking businesses
                await ListBookingBusinessesAsync(graphClient);

                // Test 2: Get specific business and its details
                Console.Write("\nEnter Business ID to get details (or press Enter to skip): ");
                var businessId = Console.ReadLine();

                if (!string.IsNullOrWhiteSpace(businessId))
                {
                    await GetBusinessDetailsAsync(graphClient, businessId);
                    await ListAppointmentsAsync(graphClient, businessId);
                    await ListServicesAsync(graphClient, businessId);
                    await ListStaffMembersAsync(graphClient, businessId);
                }

                Console.WriteLine("\n=== Complete ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ Error: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        /// <summary>
        /// Create and configure Microsoft Graph client
        /// </summary>
        private static GraphServiceClient CreateGraphClient()
        {
            // For app-only authentication (daemon/service scenarios)
            var clientSecretCredential = new ClientSecretCredential(
                TenantId,
                ClientId,
                ClientSecret
            );

            // Alternative: For delegated authentication (user context)
            // var interactiveBrowserCredential = new InteractiveBrowserCredential(
            //     new InteractiveBrowserCredentialOptions
            //     {
            //         TenantId = TenantId,
            //         ClientId = ClientId,
            //         RedirectUri = new Uri("http://localhost")
            //     }
            // );

            var graphClient = new GraphServiceClient(clientSecretCredential, new[]
            {
                "https://graph.microsoft.com/.default"
            });

            Console.WriteLine("✓ Graph client initialized\n");
            return graphClient;
        }

        /// <summary>
        /// List all booking businesses in the tenant
        /// </summary>
        private static async Task ListBookingBusinessesAsync(GraphServiceClient graphClient)
        {
            Console.WriteLine("📋 Listing all booking businesses...\n");

            var businesses = await graphClient.Solutions.BookingBusinesses
                .GetAsync();

            if (businesses?.Value == null || businesses.Value.Count == 0)
            {
                Console.WriteLine("   No booking businesses found");
                return;
            }

            Console.WriteLine($"   Found {businesses.Value.Count} business(es):\n");

            foreach (var business in businesses.Value)
            {
                Console.WriteLine($"   • {business.DisplayName}");
                Console.WriteLine($"     ID: {business.Id}");
                Console.WriteLine($"     Email: {business.Email}");
                Console.WriteLine($"     Phone: {business.Phone}");
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Get details about a specific booking business
        /// </summary>
        private static async Task GetBusinessDetailsAsync(GraphServiceClient graphClient, string businessId)
        {
            Console.WriteLine($"\n🏢 Getting business details for: {businessId}\n");

            var business = await graphClient.Solutions.BookingBusinesses[businessId]
                .GetAsync();

            if (business == null)
            {
                Console.WriteLine("   Business not found");
                return;
            }

            Console.WriteLine($"   Display Name: {business.DisplayName}");
            Console.WriteLine($"   Email: {business.Email}");
            Console.WriteLine($"   Phone: {business.Phone}");
            Console.WriteLine($"   Website: {business.WebSiteUrl}");
            Console.WriteLine($"   Address: {business.Address?.Street}, {business.Address?.City}");
            Console.WriteLine($"   Default Currency: {business.DefaultCurrencyIso}");
            Console.WriteLine($"   Is Published: {business.IsPublished}");
            Console.WriteLine($"   Public URL: {business.PublicUrl}");
        }

        /// <summary>
        /// List appointments for a booking business
        /// </summary>
        private static async Task ListAppointmentsAsync(GraphServiceClient graphClient, string businessId)
        {
            Console.WriteLine($"\n📅 Listing appointments...\n");

            try
            {
                var appointments = await graphClient.Solutions.BookingBusinesses[businessId]
                    .Appointments
                    .GetAsync();

                if (appointments?.Value == null || appointments.Value.Count == 0)
                {
                    Console.WriteLine("   No appointments found");
                    return;
                }

                Console.WriteLine($"   Found {appointments.Value.Count} appointment(s):\n");

                foreach (var appointment in appointments.Value)
                {
                    Console.WriteLine($"   • Appointment ID: {appointment.Id}");
                    Console.WriteLine($"     Customer: {appointment.CustomerName}");
                    Console.WriteLine($"     Email: {appointment.CustomerEmailAddress}");
                    Console.WriteLine($"     Start: {appointment.StartDateTime}");
                    Console.WriteLine($"     End: {appointment.EndDateTime}");
                    Console.WriteLine($"     Duration: {appointment.Duration}");
                    Console.WriteLine($"     Price: {appointment.Price} {appointment.PriceType}");
                    Console.WriteLine($"     Online: {appointment.IsLocationOnline}");
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ❌ Error listing appointments: {ex.Message}");
            }
        }

        /// <summary>
        /// List services offered by a booking business
        /// </summary>
        private static async Task ListServicesAsync(GraphServiceClient graphClient, string businessId)
        {
            Console.WriteLine($"\n🛎️  Listing services...\n");

            try
            {
                var services = await graphClient.Solutions.BookingBusinesses[businessId]
                    .Services
                    .GetAsync();

                if (services?.Value == null || services.Value.Count == 0)
                {
                    Console.WriteLine("   No services found");
                    return;
                }

                Console.WriteLine($"   Found {services.Value.Count} service(s):\n");

                foreach (var service in services.Value)
                {
                    Console.WriteLine($"   • {service.DisplayName}");
                    Console.WriteLine($"     ID: {service.Id}");
                    Console.WriteLine($"     Duration: {service.DefaultDuration}");
                    Console.WriteLine($"     Price: {service.DefaultPrice}");
                    Console.WriteLine($"     Description: {service.Description}");
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ❌ Error listing services: {ex.Message}");
            }
        }

        /// <summary>
        /// List staff members of a booking business
        /// </summary>
        private static async Task ListStaffMembersAsync(GraphServiceClient graphClient, string businessId)
        {
            Console.WriteLine($"\n👥 Listing staff members...\n");

            try
            {
                var staffMembers = await graphClient.Solutions.BookingBusinesses[businessId]
                    .StaffMembers
                    .GetAsync();

                if (staffMembers?.Value == null || staffMembers.Value.Count == 0)
                {
                    Console.WriteLine("   No staff members found");
                    return;
                }

                Console.WriteLine($"   Found {staffMembers.Value.Count} staff member(s):\n");

                foreach (var staff in staffMembers.Value)
                {
                    Console.WriteLine($"   • {staff.DisplayName}");
                    Console.WriteLine($"     ID: {staff.Id}");
                    Console.WriteLine($"     Email: {staff.EmailAddress}");
                    Console.WriteLine($"     Role: {staff.Role}");
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ❌ Error listing staff: {ex.Message}");
            }
        }

        /// <summary>
        /// Example: Create a new appointment
        /// </summary>
        private static async Task<BookingAppointment> CreateAppointmentAsync(
            GraphServiceClient graphClient,
            string businessId,
            string customerName,
            string customerEmail,
            string serviceId,
            DateTime startTime,
            DateTime endTime)
        {
            Console.WriteLine($"\n📝 Creating new appointment...\n");

            var newAppointment = new BookingAppointment
            {
                CustomerName = customerName,
                CustomerEmailAddress = customerEmail,
                ServiceId = serviceId,
                StartDateTime = new DateTimeTimeZone
                {
                    DateTime = startTime.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = "UTC"
                },
                EndDateTime = new DateTimeTimeZone
                {
                    DateTime = endTime.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = "UTC"
                },
                IsLocationOnline = true,
                SmsNotificationsEnabled = true,
                IsCustomerAllowedToManageBooking = true,
                CustomerTimeZone = "Europe/Dublin"
            };

            var appointment = await graphClient.Solutions.BookingBusinesses[businessId]
                .Appointments
                .PostAsync(newAppointment);

            Console.WriteLine($"   ✓ Appointment created: {appointment.Id}");
            return appointment;
        }
    }
}
