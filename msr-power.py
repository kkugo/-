
import datetime
from pymodbus.client.sync import ModbusSerialClient as ModbusClient

fname='rec\\power'+str(datetime.date.today())+'.txt'


client = ModbusClient(method = "rtu", port="COM3",stopbits = 2, bytesize = 8, parity = 'N', baudrate= 9600, timeout= 1)
client.connect()

result= client.read_holding_registers(address=0x200, count=10, unit= 1)

wh=65536*result.registers[0]+result.registers[1]
vah=65536*result.registers[8]+result.registers[9]

client.close() 

with open(fname,'a') as  f:
  st=str(datetime.datetime.now())+' '
  st=st+str(wh)+' '
  st=st+str(vah)
  print(st,file=f)


