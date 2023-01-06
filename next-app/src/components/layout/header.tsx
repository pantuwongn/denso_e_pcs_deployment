import { FC } from 'react'

interface IProps {
  title: string
}

const Header: FC<IProps> = ({ title }: IProps) => {

  return (
    <>
      <div className="w-full flex justify-between items-center p-4 shadow shadow-zinc-500">
        <div className="flex items-center">
          <div className="text-xl font-bold mr-4">{title}</div>
        </div>
      </div>
    </>
  )
}

export default Header