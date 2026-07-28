using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using QYachtMaster.Database;
using QYachtMaster.MVVM;
using QYachtMaster.Utils;

namespace QYachtMaster.ViewModels
{
    public class BoatCard : ViewModelBase
    {
        private string _status = "In Progress";
        private string _bayName = "Bay 1: Mechanical";
        private string _technicianName = string.Empty;

        public int JobCardID { get; set; }
        public string JobCardNo { get; set; } = string.Empty;
        public string CustomerName { get; set; } = string.Empty;
        public string MobilePhone { get; set; } = string.Empty;
        public string RegNo { get; set; } = string.Empty;
        public string MakeModel { get; set; } = string.Empty;
        public string JobDate { get; set; } = string.Empty;

        public string Status
        {
            get => _status;
            set => SetProperty(ref _status, value);
        }

        public string BayName
        {
            get => _bayName;
            set => SetProperty(ref _bayName, value);
        }

        public string TechnicianName
        {
            get => _technicianName;
            set => SetProperty(ref _technicianName, value);
        }
    }

    public class ServiceBayViewModel : ViewModelBase
    {
        public string BayName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string HeaderColor { get; set; } = "#1E293B";

        public ObservableCollection<BoatCard> BoatCards { get; set; } = new();
    }

    public class DispatcherBoardViewModel : ViewModelBase
    {
        private bool _isLoading = false;

        public ObservableCollection<ServiceBayViewModel> Bays { get; set; } = new();

        public bool IsLoading
        {
            get => _isLoading;
            set => SetProperty(ref _isLoading, value);
        }

        public ICommand RefreshCommand { get; }

        public DispatcherBoardViewModel()
        {
            RefreshCommand = new RelayCommand(async _ => await LoadDataAsync());
            InitializeBays();
            LanSyncEngine.MessageReceived += OnLanSyncReceived;
            Task.Run(async () => await LoadDataAsync());
        }

        private void InitializeBays()
        {
            Bays.Clear();
            Bays.Add(new ServiceBayViewModel { BayName = "Bay 1: Mechanical / المحركات", Description = "Mechanical Repairs & Engine Servicing", HeaderColor = "#3B82F6" });
            Bays.Add(new ServiceBayViewModel { BayName = "Bay 2: Electrical / الكهرباء", Description = "Avionics, Wiring & Marine Electronics", HeaderColor = "#8B5CF6" });
            Bays.Add(new ServiceBayViewModel { BayName = "Bay 3: Hull Repair / الهيكل", Description = "Fiberglass, Gelcoat & Anti-Fouling", HeaderColor = "#F59E0B" });
            Bays.Add(new ServiceBayViewModel { BayName = "Dock 4: Inspection / الفحص النهائي", Description = "Sea Trial, Wash & Quality Control", HeaderColor = "#10B981" });
        }

        public async Task LoadDataAsync()
        {
            IsLoading = true;
            try
            {
                string query = @"
                    SELECT j.JobCardID, j.JobCardNo, j.Status, j.JobDate,
                           c.Name AS CustName, c.MobilePhone,
                           v.RegNo, v.Make, v.Model
                    FROM JobCards j
                    JOIN Customers c ON j.CustomerID = c.CustomerID
                    JOIN Vehicles v ON j.VehicleID = v.VehicleID
                    WHERE j.Status != 'Completed' AND j.Status != 'Cancelled'
                    ORDER BY j.JobCardID DESC LIMIT 1000;";

                var cards = await Task.Run(() => DatabaseHelper.ExecuteQuery(query, reader =>
                {
                    string make = reader["Make"]?.ToString() ?? "";
                    string model = reader["Model"]?.ToString() ?? "";
                    string makeModel = $"{make} {model}".Trim();
                    if (string.IsNullOrEmpty(makeModel)) makeModel = "Marine Vessel";

                    return new BoatCard
                    {
                        JobCardID = Convert.ToInt32(reader["JobCardID"]),
                        JobCardNo = reader["JobCardNo"].ToString()!,
                        Status = reader["Status"].ToString()!,
                        JobDate = reader["JobDate"].ToString()!,
                        CustomerName = reader["CustName"].ToString()!,
                        MobilePhone = reader["MobilePhone"].ToString()!,
                        RegNo = reader["RegNo"]?.ToString() ?? "N/A",
                        MakeModel = makeModel
                    };
                }));

                App.Current.Dispatcher.Invoke(() =>
                {
                    foreach (var bay in Bays) bay.BoatCards.Clear();

                    int index = 0;
                    foreach (var card in cards)
                    {
                        // Map status/index to bay
                        int bayIdx = index % 4;
                        card.BayName = Bays[bayIdx].BayName;
                        Bays[bayIdx].BoatCards.Add(card);
                        index++;
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DispatcherBoard Load Error: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task MoveBoatToBayAsync(BoatCard card, ServiceBayViewModel targetBay)
        {
            if (card == null || targetBay == null) return;

            var oldBay = Bays.FirstOrDefault(b => b.BoatCards.Contains(card));
            if (oldBay == targetBay) return;

            oldBay?.BoatCards.Remove(card);
            card.BayName = targetBay.BayName;
            targetBay.BoatCards.Add(card);

            // Update in DB
            string newStatus = targetBay.BayName.Contains("Dock 4") ? "Ready for Pickup" : "In Progress";
            card.Status = newStatus;

            await Task.Run(() =>
            {
                string sql = "UPDATE JobCards SET Status = @s WHERE JobCardID = @id;";
                var p1 = new Microsoft.Data.Sqlite.SqliteParameter("@s", newStatus);
                var p2 = new Microsoft.Data.Sqlite.SqliteParameter("@id", card.JobCardID);
                DatabaseHelper.ExecuteNonQuery(sql, p1, p2);
            });

            // Broadcast real-time LAN update across workstations
            LanSyncEngine.BroadcastEvent("BAY_CHANGED", card.JobCardID, targetBay.BayName);
        }

        private void OnLanSyncReceived(LanSyncMessage msg)
        {
            if (msg.EventType == "BAY_CHANGED" || msg.EventType == "JOBCARD_UPDATED")
            {
                App.Current.Dispatcher.Invoke(async () => await LoadDataAsync());
            }
        }
    }
}
