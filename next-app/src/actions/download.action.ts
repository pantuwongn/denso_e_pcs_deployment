import axiosInstance from "@/lib/axios";
import { FormJsonResponse, PCSUploadForm } from "@/types/upload.type";

export async function getFormJson(): Promise<FormJsonResponse> {
  const { data } = await axiosInstance.get<FormJsonResponse>('/get_mock_data')
  return data
}

export async function convertJsonToXlsx(uploadForm: PCSUploadForm): Promise<Blob> {
  const { data } = await axiosInstance.post<Blob>('/convert_json_to_xlsx', uploadForm, {
    responseType: 'blob'
  })
  return data
}
