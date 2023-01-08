export interface DownloadResponse {

}

export interface FormJsonResponse extends PCSUploadForm {

}

export interface PCSUploadForm {
  pcs_no: string
  date: string
  status: string
  line: string
  assy_name: string
  part_name: string
  customer: string
  processes: PCSProcess
}

export interface PCSProcess {
  name: string
  items: PCSProcessItem[]
}

export interface PCSProcessItem {
  control_item_no: string
  control_item_type: string
  parameter: ItemParameter
  sc_symbols: ItemSCSymbol[]
  check_timing: string
  control_method: ItemControlMethod
  initial_p_capability: ItemCapability
  remark: ItemRemark
  measurement: string
  readability: string
  start_effective: string
}

export interface ItemParameter {
  parameter: string
  master_value: number
  limit_type: string
  symbolic: string
  limit: string
  sign: string
  tolerance_up: string
  tolerance_down: string
  upper_limit: string
  lower_limit: string
  unit: string
}

export interface ItemSCSymbol {
  character: string
  shape: string
}

export interface ItemControlMethod {
  sample_no: string
  interval: string
  "100_method": string
  in_charge: string
  calibration_interval: string
}

export interface ItemCapability {
  x_bar: string
  cpk: string
}

export interface ItemRemark {
  remark: string
  ws_no: string
  related_std: string
}
