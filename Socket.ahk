CreateServerSocket(port)
{
	; DllCall("Ws2_32\WSAStartup", "UShort", 0x0202, "Ptr", 0)
	VarSetCapacity(wsaData, 32)
	DllCall("Ws2_32\WSAStartup", "UShort", 2, "UInt", &wsaData)
	socket := DllCall("Ws2_32\socket", "Int", 2, "Int", 1, "Int", 0)
	if (socket = -1)
		return -1

	VarSetCapacity(addr, 16, 0)
	NumPut(2, addr, 0, "UShort")
	NumPut(DllCall("Ws2_32\htons", "UShort", port), addr, 2, "UShort")
	NumPut(DllCall("Ws2_32\inet_addr", "AStr", "0.0.0.0"), addr, 4, "UInt")

	if (DllCall("Ws2_32\bind", "UInt", socket, "Ptr", &addr, "Int", 16) != 0)
	{
		DllCall("Ws2_32\closesocket", "UInt", socket)
		return -1
	}

	if (DllCall("Ws2_32\listen", "UInt", socket, "Int", 5) != 0)
	{
		DllCall("Ws2_32\closesocket", "UInt", socket)
		return -1
	}

	return socket
}
