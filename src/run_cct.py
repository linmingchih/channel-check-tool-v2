import os
import re
import uuid
from ansys.aedt.core import Circuit
from ansys.aedt.core.generic.constants import Setups
import numpy as np
wkdir = 'd:/demo'


def integrate_nonuniform(x_list, y_list):
    integral = 0.0
    for i in range(len(x_list) - 1):
        # 使用梯形公式
        dx = x_list[i + 1] - x_list[i]
        integral += 0.5 * (y_list[i] + y_list[i + 1]) * dx
    return integral


def get_sig_isi(time_list, voltage_list, unit_interval):
    """
    sig: 寬度為 unit_interval 的滑動視窗內，∫v(t)dt 的最大值（視窗完全落在資料範圍內）
    isi: 除了該最大視窗以外，∫|v(t)|dt
    """
    t = np.asarray(time_list, dtype=float)
    v = np.asarray(voltage_list, dtype=float)
    if t.ndim != 1 or v.ndim != 1 or t.size != v.size:
        raise ValueError("time_list 與 voltage_list 必須為一維且長度相同")
    if unit_interval <= 0:
        raise ValueError("unit_interval 必須為正數")

    # 依時間排序
    order = np.argsort(t)
    t = t[order]
    v = v[order]

    if t[-1] - t[0] < unit_interval:
        raise ValueError("資料時間範圍小於 unit_interval，無法形成完整視窗")

    # 累積積分（梯形法）
    dt = np.diff(t)
    trap = np.concatenate([[0.0], np.cumsum((v[:-1] + v[1:]) * 0.5 * dt)])
    trap_abs = np.concatenate([[0.0], np.cumsum((np.abs(v[:-1]) + np.abs(v[1:])) * 0.5 * dt)])
    total_abs = trap_abs[-1]

    # 僅遍歷「能形成完整視窗」的起點：t[i] + UI <= t[-1]
    n = len(t)
    # 最後可用起點的索引上限（滿足 t[i] + UI <= t[-1]）
    last_i = np.searchsorted(t, t[-1] - unit_interval, side="right") - 1
    if last_i < 0:
        raise ValueError("沒有任何起點能形成完整視窗")

    sig_max = -np.inf
    best_i = best_j = 0
    best_t_end = None
    j = 0

    for i in range(last_i + 1):
        t_end = t[i] + unit_interval

        # 向右移動 j，使得 t[j] <= t_end < t[j+1]（或 j 到尾）
        while j + 1 < n and t[j + 1] <= t_end:
            j += 1

        # 此時 t_end 一定 <= t[-1]，不會超界
        integ = trap[j] - trap[i]
        # 若 t_end 在 (t[j], t[j+1])，補最後一段梯形
        if j + 1 < n and t[j] < t_end < t[j + 1]:
            v_end = v[j] + (v[j + 1] - v[j]) * (t_end - t[j]) / (t[j + 1] - t[j])
            integ += 0.5 * (v[j] + v_end) * (t_end - t[j])

        if integ > sig_max:
            sig_max = integ
            best_i, best_j, best_t_end = i, j, t_end

    # 以最佳視窗計算 |v| 的積分，供 isi 使用
    i, j, t_end = best_i, best_j, best_t_end
    integ_abs = (trap_abs[j] - trap_abs[i])
    if j + 1 < n and t[j] < t_end < t[j + 1]:
        v_end = v[j] + (v[j + 1] - v[j]) * (t_end - t[j]) / (t[j + 1] - t[j])
        integ_abs += 0.5 * (abs(v[j]) + abs(v_end)) * (t_end - t[j])

    sig = float(sig_max)
    isi = float(total_abs - integ_abs)
    return sig, isi




class Tx:
    def __init__(self, pid, vhigh, t_rise, ui, res_tx, cap_tx):
        self.pid = pid
        self.active = [f"V{pid} netb_{pid} 0 PULSE(0 {vhigh} 1e-10 {t_rise} {t_rise} {ui} 1.5e+100)",
                       f"R{pid} netb_{pid} net_{pid} {res_tx}",
                       f"C{pid} netb_{pid} 0 {cap_tx}"]
        self.passive = [f"R{pid} netb_{pid} net_{pid} {res_tx}",
                        f"C{pid} netb_{pid} 0 {cap_tx}"]

    def get_netlist(self, active=True):
        if active:
            return self.active
        else:
            return self.passive
        

class Tx_diff:
    def __init__(self, pid_pos, pid_neg, vhigh, t_rise, ui, res_tx, cap_tx):
        self.pid_pos = pid_pos
        self.pid_neg = pid_neg
        
        self.active = [f"V{pid_pos} netb_{pid_pos} 0 PULSE(0 {vhigh} 1e-10 {t_rise} {t_rise} {ui} 1.5e+100)",
                       f"R{pid_pos} netb_{pid_pos} net_{pid_pos} {res_tx}",
                       f"C{pid_pos} netb_{pid_pos} 0 {cap_tx}",
                       f"V{pid_neg} netb_{pid_neg} 0 PULSE(0 -{vhigh} 1e-10 {t_rise} {t_rise} {ui} 1.5e+100)",
                       f"R{pid_neg} netb_{pid_neg} net_{pid_neg} {res_tx}",
                       f"C{pid_neg} netb_{pid_neg} 0 {cap_tx}"]
        
        self.passive = [f"R{pid_pos} netb_{pid_pos} net_{pid_pos} {res_tx}",
                        f"C{pid_pos} netb_{pid_pos} 0 {cap_tx}",
                        f"R{pid_neg} netb_{pid_neg} net_{pid_neg} {res_tx}",
                        f"C{pid_neg} netb_{pid_neg} 0 {cap_tx}"]

    def get_netlist(self, active=True):
        if active:
            return self.active
        else:
            return self.passive
        

        
    
class Rx:
    def __init__(self, pid, res_rx, cap_rx):
        self.pid = pid
        self.netlist = [f'R{pid} net_{pid} 0 {res_rx}', 
                        f'C{pid} net_{pid} 0 {cap_rx}']
        self.waveforms = {}
        
    def get_netlist(self):
        return self.netlist
    
class Rx_diff:
    def __init__(self, pid_pos, pid_neg, res_rx, cap_rx):
        self.pid_pos = pid_pos
        self.pid_neg = pid_neg
        
        self.netlist = [f'R{pid_pos} net_{pid_pos} 0 {res_rx}', 
                        f'C{pid_pos} net_{pid_pos} 0 {cap_rx}',
                        f'R{pid_neg} net_{pid_neg} 0 {res_rx}', 
                        f'C{pid_neg} net_{pid_neg} 0 {cap_rx}',]
        self.waveforms = {}
        
    def get_netlist(self):
        return self.netlist    
    
    
    
class Design:
    def __init__(self, tstep='100ps', tstop='3ns'):
        self.netlist_path = os.path.join(wkdir, f"{uuid.uuid4()}.cir")
        open(self.netlist_path, 'w').close()
        
        self.circuit = circuit = Circuit(version='2025.1', 
                                         non_graphical=True,
                                         close_on_exit=True)
    
        circuit.add_netlist_datablock(self.netlist_path)
        self.setup = circuit.create_setup('myTransient', Setups.NexximTransient)
        self.setup.props['TransientData'] = [tstep, tstop]
        self.circuit.save_project()

    def run(self, netlist):
        with open(self.netlist_path, 'w') as f:
            f.write(netlist)
            
        self.circuit.odesign.InvalidateSolution('myTransient')
        self.circuit.save_project()
        self.circuit.analyze('myTransient')
        self.circuit.save_project()
        
        result = {}
        for v in self.circuit.post.available_report_quantities():
            data = self.circuit.post.get_solution_data(v, domain='Time')
            x = [1e3*i for i in data.primary_sweep_values]
            y = [1e-3*i for i in data.data_real()]
            m = re.search(r'net_(\d+)', v)
            if m:
                number = int(m.group(1))
                result[number] = (x, y)
        return result
    
class CCT:
    def __init__(self, snp_path, tx_ports, rx_ports, tx_diff_ports, rx_diff_ports):
        self.snp_path = snp_path
        
        self.tx_ports = tx_ports
        self.rx_ports = rx_ports
        self.tx_diff_ports = tx_diff_ports
        self.rx_diff_ports = rx_diff_ports 
        
        self.txs = []
        self.rxs = []
        
        number = int(os.path.basename(snp_path).split('.')[-1][1:-1])
        nets = ' '.join([f'net_{i+1}' for i in range(number)])
        self.netlist= [f'.model "Channel" S TSTONEFILE="{snp_path}" INTERPOLATION=LINEAR INTDATTYP=MA HIGHPASS=10 LOWPASS=10 convolution=1 enforce_passivity=0 Noisemodel=External',
                       f'S1 {nets} FQMODEL="Channel"']
    
    def set_txs(self, vhigh, t_rise, ui, res_tx, cap_tx):
        self.ui = ui
        for pid in self.tx_ports:
            self.txs.append(Tx(pid, vhigh, t_rise, ui, res_tx, cap_tx))
        
        for pid_pos, pid_neg in self.tx_diff_ports:
            self.txs.append(Tx_diff(pid_pos, pid_neg, vhigh, t_rise, ui, res_tx, cap_tx))
            
    def set_rxs(self, res_rx, cap_rx):
        for pid in self.rx_ports:
            self.rxs.append(Rx(pid, res_rx, cap_rx))

        for pid_pos, pid_neg in self.rx_diff_ports:
            self.rxs.append(Rx_diff(pid_pos, pid_neg, res_rx, cap_rx))

    
    
    def run(self, tstep='100ps', tstop='3ns'):
        design = Design(tstep, tstop)
        for tx1 in self.txs:
            netlist = [i for i in self.netlist]
            for tx2 in self.txs:
                if tx2 == tx1:
                    netlist += tx2.get_netlist(True)
                else:
                    netlist += tx2.get_netlist(False)
            
            for rx in self.rxs:
                netlist += rx.get_netlist()
  
            result = design.run('\n'.join(netlist))
            
            for rx in self.rxs:
                if type(rx) == Rx:
                    rx.waveforms[tx1] = result[rx.pid]
                
                if type(rx) == Rx_diff:
                    time, waveform_pos = result[rx.pid_pos]
                    time, waveform_neg = result[rx.pid_neg]
                    new_result = (time, [vpos-vneg for vpos, vneg in zip(waveform_pos, waveform_neg)])
                    rx.waveforms[tx1] = new_result
            
    def calculate(self, output_path):
        ui = float(self.ui.replace('ps', '')) 
        
        for rx in self.rxs:
            data = []
            for tx, (time, voltage) in rx.waveforms.items():
                data.append((max(voltage), tx))
            
            peak, tx = sorted(data)[-1]
            rx.tx = tx
        
        result = []
        for rx in self.rxs:
            xtalk = 0
            for tx, waveform in rx.waveforms.items():
                time, voltage = waveform 
                if tx == rx.tx: 
                    sig, isi = get_sig_isi(time, voltage, ui)
                    continue
                
                xtalk += integrate_nonuniform(time, [abs(v) for v in voltage])
            pseudo_eye = sig - isi - xtalk
            p_ratio = sig / (isi + xtalk)
            
            if type(rx) == Rx:
                tx_id = rx.tx.pid
                rx_id = rx.pid
            elif type(rx) == Rx_diff: 
                tx_id = f'{rx.tx.pid_pos}_{rx.tx.pid_neg}'
                rx_id = f'{rx.pid_pos}_{rx.pid_neg}'
            
            result.append(f'{tx_id:5}, {rx_id:5}, {sig:10.3f}, {isi:10.3f}, {xtalk:10.3f}, {pseudo_eye:10.3f}, {p_ratio:10.3f}')
            
        with open(output_path, 'w') as f:
            f.writelines('tx_id, rx_id, sig(V*ps), isi(V*ps), xtalk(V*ps), pseudo_eye(V*ps), power_ratio\n')
            f.write('\n'.join(result))                

    
    
if __name__ == '__main__':
    touchstone_path = r"D:\OneDrive - ANSYS, Inc\a-client-repositories\quanta-cct-circuit-202508\data\channels.s12p"
    cct = CCT(touchstone_path, 
              tx_ports=[5,6], 
              rx_ports=[11,12],
              tx_diff_ports = [(1, 2), (3, 4)],
              rx_diff_ports = [(7, 8), (9, 10)])
    
    cct.set_txs(vhigh="0.8V", t_rise="30ps", ui="133ps", res_tx="40ohm", cap_tx="1pF")
    cct.set_rxs(res_rx="30ohm", cap_rx="1.8pF")
    cct.run()
    cct.calculate(output_path='d:/demo/cct.csv')
    