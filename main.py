from netmiko import ConnectHandler
import paramiko
import configparser
import json
import datetime
import xlsxwriter

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
    # 設定ファイルの読み込み
    config = configparser.ConfigParser()
    config.read('config.ini')

    # 踏み台の情報
    jump_host_username = config['credentials']['jump_host_username']
    jump_host_password = config['credentials']['jump_host_password']

    # 各ホストの情報
    hosts = {}
    for host in config['credentials']:
        if host.startswith('host') and host != 'host':
            hostname = host.split('_')[0]
            username = config['credentials'][f'{hostname}_username']
            password = config['credentials'][f'{hostname}_password']
            ip = config['credentials'][f'{hostname}_ip']
            hosts[hostname] = {'username': username, 'password': password, 'ip': ip}

    # 踏み台経由で各ホストにアクセスし、データを取得してExcelに書き込む
    for hostname, host_info in hosts.items():
        print(f"Processing host: {hostname}")
        with ConnectHandler(
            device_type='juniper_junos',
            host=host_info['ip'],
            username=jump_host_username,
            password=jump_host_password,
            sock=paramiko.ProxyCommand(f'ssh -o StrictHostKeyChecking=no -W %h:%p {jump_host_username}@{jump_host_ip}')
        ) as ssh:
            output = ssh.send_command('show interface extensive | display json')

        interface_data = extract_interface_data(output)
        write_to_excel(interface_data, f'{hostname}_interface_stats.xlsx', hostname)
        print(f"Completed processing host: {hostname}")

if __name__ == "__main__":
    main()
