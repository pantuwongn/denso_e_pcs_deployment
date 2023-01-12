import { Spin } from 'antd'
import { FC } from 'react'
import Header from './header'
import { LayoutStore } from '@/store/layout.store'

interface IProps {
  title: string
  children?: React.ReactNode
}

const Container: FC<IProps> = ({ title, children }: IProps) => {
  const { isLoading } = LayoutStore()
  return (
    <Spin spinning={isLoading} size="large" style={{ top: '50%', transform: 'translateY(-50%)' }}>
      <Header title={title}/>
      <div className="w-full overflow-y-auto overflow-x-hidden" style={{ height: "calc(100vh - 64px)" }}>
        <div className="w-full h-screen p-4 m-auto">
          {children}
        </div>
      </div>
    </Spin>
  )
}

export default Container