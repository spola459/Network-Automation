from netmiko import ConnectHandler
from openpyxl import load_workbook, Workbook

# Switch connection details
device = {
    "device_type": "cisco_ios",
    "ip": "core_IP",      # Core switch mgmt IP
    "username": "admin",
    "password": "password",
    "secret": "password",    # enable password (if required)
}

# Input and output Excel files
input_file = "input_ips.xlsx"   # Must have IPs in column A
output_file = "output_mac.xlsx"

def get_mac_from_ip(connection, target_ip):
    # Step 1: Get ARP entry for IP
    output = connection.send_command(f"show ip arp {target_ip}")
    if "Incomplete" in output or target_ip not in output:
        return None, None

    mac_address = None
    vlan_id = None

    for line in output.splitlines():
        if target_ip in line:
            parts = line.split()
            # Expected: [Protocol, Address, Age, MAC, Type, Interface]
            if len(parts) >= 6:
                mac_address = parts[3]   # <-- Correct MAC position
                vlan_id = parts[-1]      # Interface (e.g., Vlan10)
            break

    if not mac_address:
        return None, None

    # Step 2: Lookup MAC in switch MAC table
    mac_output = connection.send_command(f"show mac address-table | include {mac_address}")
    interface = None
    if mac_output.strip():
        for line in mac_output.splitlines():
            if mac_address in line:
                parts = line.split()
                if len(parts) >= 4:
                    interface = parts[-1]  # last column is interface
                break

    return mac_address, interface

def main():
    # Connect to switch
    connection = ConnectHandler(**device)
    connection.enable()

    # Load IP list from Excel
    wb = load_workbook(input_file)
    ws = wb.active
    ip_list = [cell.value for cell in ws['A'] if cell.value]

    # Create output Excel
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.append(["IP Address", "MAC Address", "Interface"])

    # Process each IP
    for ip in ip_list:
        print(f"[*] Checking IP: {ip}")
        mac, interface = get_mac_from_ip(connection, ip)
        if mac:
            print(f"[+] {ip} → {mac} ({interface})")
        else:
            print(f"[!] {ip} → No entry found")
        out_ws.append([ip, mac if mac else "Not Found", interface if interface else "N/A"])

    # Save results
    out_wb.save(output_file)
    print(f"\n[✓] Results saved to {output_file}")

    connection.disconnect()


if __name__ == "__main__":

    main()

