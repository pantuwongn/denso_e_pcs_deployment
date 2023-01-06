import { convertJsonToXlsx, getFormJson } from "@/actions/download.action";
import { LayoutStore } from "@/store/layout.store";
import { DownloadBlob } from "@/util";
import { Button } from "antd";
import { FC } from "react";

interface IProps {
}

const DownloadXlsxView: FC<IProps> = ({ }: IProps) => {
  const { setIsLoading } = LayoutStore()
  const downloadXlsx = async () => {
    setIsLoading(true)
    const pcsForm = await getFormJson()
    const blob = await convertJsonToXlsx(pcsForm)
    const currentDate = new Date()
    DownloadBlob(`e-pcs-${currentDate.getFullYear()}-${currentDate.getMonth()+1}-${currentDate.getDate()}`,blob)
    setIsLoading(false)
  }

  return (
    <div className='w-full h-full flex flex-col justify-center items-center'>
      <Button type="primary" onClick={downloadXlsx}>Download XLSX</Button>
    </div>
  )
}

export default DownloadXlsxView;
