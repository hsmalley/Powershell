$mac = [byte[]](0x00, 0x25, 0x64, 0x79, 0x24, 0xeb)
$UDPclient = new-Object System.Net.Sockets.UdpClient
$UDPclient.Connect(([System.Net.IPAddress]::Broadcast),4000)
$packet = [byte[]](,0xFF * 102)
6..101 |% { $packet[$_] = $mac[($_%6)]}
"Send: "
$packet
$UDPclient.Send($packet, $packet.Length)
