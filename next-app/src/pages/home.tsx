import Head from 'next/head'
import type { NextPage } from 'next'
import Container from '@/components/layout'
import DownloadXlsxView from '@/views/download-xlsx-view'

const Home: NextPage = () => {
  return (
    <>
      <div>
        <Head>
          <title>e-PCS Control item generator</title>
          <link rel="icon" href="/favicon.ico" />
        </Head>
      </div>
      <Container title='e-PCS Control item generator'>
        <DownloadXlsxView />
      </Container>
    </>
  )
}

export default Home
