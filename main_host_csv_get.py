from netmiko import ConnectHandler
import paramiko
import json
import datetime
import xlsxwriter
import csv

def extract_interface_data(json_output):
    interface_data = []
    interfaces = json.loads(json_output)['configuration']['interfaces']['interface']
    for interface in interfaces:
        name = interface['name']
        multicast_in = interface['traffic-statistics']['multicast-packets']['input']
        multicast_out = interface['traffic-statistics']['multicast-packets']['output']
        broadcast_in = interface['traffic-statistics']['broadcast-packets']['input']
        broadcast_out = interface['traffic-statistics']['broadcast-packets']['output']
        interface_data.append((name, multicast_in, multicast_out, broadcast_in, broadcast_out))
    return interface_data

def write_to_excel(data, filename, hostname):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Write headers
    headers = ['Date', 'Hostname', 'Multicast (RX)', 'Multicast (TX)', 'Broadcast (RX)', 'Broadcast (TX)']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # Write data
    for row, (name, multicast_in, multicast_out, broadcast_in, broadcast_out) in enumerate(data, start=1):
        worksheet.write(row, 0, datetime.date.today().strftime('%Y-%m-%d'))
        worksheet.write(row, 1, hostname)
        worksheet.write(row, 2, multicast_in)
        worksheet.write(row, 3, multicast_out)
        worksheet.write(row, 4, broadcast_in)
        worksheet.write(row, 5, broadcast_out)

    workbook.close()

def main():
    # CSVファイルからホストの情報を読み取る
    hosts = []
    with open('hosts.csv', newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            hostname = row.get('hostname')
            ip = row.get('ip')
            if not hostname and not ip:
                continue

            hosts.append({'hostname': hostname, 'ip': ip})

    # 踏み台経由で各ホストにアクセスし、データを取得してExcelに書き込む
    for host_info in hosts:
        hostname = host_info['hostname']
        ip = host_info['ip']
        print(f"Processing host: {hostname or ip}")
        with ConnectHandler(
            device_type='juniper_junos',
            host=ip,
            username='your_username',  # ここにユーザ名を入力
            password='your_password',  # ここにパスワードを入力
            sock=paramiko.ProxyCommand('ssh -o StrictHostKeyChecking=no -W %h:%p user@jump_host_ip')
        ) as ssh:
            output = ssh.send_command('show interface extensive | display json')

        interface_data = extract_interface_data(output)
        write_to_excel(interface_data, f'{hostname or ip}_interface_stats.xlsx', hostname or ip)
        print(f"Completed processing host: {hostname or ip}")

if __name__ == "__main__":
    main()
