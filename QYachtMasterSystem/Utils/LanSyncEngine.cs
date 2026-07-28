using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace QYachtMaster.Utils
{
    public class LanSyncMessage
    {
        public string EventType { get; set; } = string.Empty;
        public int EntityID { get; set; }
        public string Payload { get; set; } = string.Empty;
        public string SenderMachine { get; set; } = Environment.MachineName;
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
    }

    public static class LanSyncEngine
    {
        private const int LanSyncPort = 19840;
        private static UdpClient? _udpClient;
        private static CancellationTokenSource? _cts;
        private static bool _isRunning = false;

        public static event Action<LanSyncMessage>? MessageReceived;

        public static void StartServer()
        {
            if (_isRunning) return;
            try
            {
                _udpClient = new UdpClient();
                _udpClient.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
                _udpClient.Client.Bind(new IPEndPoint(IPAddress.Any, LanSyncPort));
                _udpClient.EnableBroadcast = true;

                _cts = new CancellationTokenSource();
                _isRunning = true;

                Task.Run(() => ListenLoop(_cts.Token));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LanSyncEngine Start failed: {ex.Message}");
            }
        }

        public static void BroadcastEvent(string eventType, int entityId = 0, string payload = "")
        {
            try
            {
                var msg = new LanSyncMessage
                {
                    EventType = eventType,
                    EntityID = entityId,
                    Payload = payload
                };

                string json = JsonSerializer.Serialize(msg);
                byte[] data = Encoding.UTF8.GetBytes(json);

                using (var sender = new UdpClient())
                {
                    sender.EnableBroadcast = true;
                    var endPoint = new IPEndPoint(IPAddress.Broadcast, LanSyncPort);
                    sender.Send(data, data.Length, endPoint);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LanSyncEngine Broadcast error: {ex.Message}");
            }
        }

        private static async Task ListenLoop(CancellationToken token)
        {
            while (!token.IsCancellationRequested && _udpClient != null)
            {
                try
                {
                    var result = await _udpClient.ReceiveAsync(token);
                    string json = Encoding.UTF8.GetString(result.Buffer);
                    var msg = JsonSerializer.Deserialize<LanSyncMessage>(json);

                    if (msg != null && msg.SenderMachine != Environment.MachineName)
                    {
                        AppDomain.CurrentDomain.SetData("LastLanEvent", msg);
                        MessageReceived?.Invoke(msg);
                    }
                }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LanSyncEngine Listen exception: {ex.Message}");
                }
            }
        }

        public static void StopServer()
        {
            _cts?.Cancel();
            _udpClient?.Close();
            _isRunning = false;
        }
    }
}
