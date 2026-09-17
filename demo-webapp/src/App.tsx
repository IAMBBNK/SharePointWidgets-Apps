import { BrowserRouter, Route, Routes } from 'react-router-dom'
import { MarketingLayout } from './layouts/MarketingLayout'
import { DemoLayout } from './layouts/DemoLayout'
import { HomePage } from './pages/HomePage'
import { ConfluenceServicePage, SharePointServicePage } from './pages/ServicePages'
import { DemoHomePage } from './pages/demo/DemoHomePage'
import { CatalogPage } from './pages/demo/CatalogPage'
import { DemoPage } from './pages/demo/DemoPage'
import { ResourcesWidget } from './widgets/ResourcesWidget'
import { TrainingsWidget } from './widgets/TrainingsWidget'
import { GovernanceWidget } from './widgets/GovernanceWidget'
import { MaintenanceFormWidget, SiteFormWidget } from './widgets/FormsWidget'
import { CompletionWidget } from './widgets/CompletionWidget'
import { KpiWidget, OtKpiWidget, PageKpiWidget } from './widgets/KpiWidgets'
import { MapWidget } from './widgets/MapWidget'
import { CarouselWidget } from './widgets/CarouselWidget'
import { TimelineWidget } from './widgets/TimelineWidget'
import { SurveyWidget } from './widgets/SurveyWidget'
import { ProjectsWidget } from './widgets/ProjectsWidget'
import { ConfluenceWidget } from './widgets/ConfluenceWidget'
import { EngageWidget } from './widgets/EngageWidget'
import { NewsletterWidget } from './widgets/NewsletterWidget'

export default function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route element={<MarketingLayout />}>
          <Route path="/" element={<HomePage />} />
          <Route path="/services/sharepoint" element={<SharePointServicePage />} />
          <Route path="/services/confluence" element={<ConfluenceServicePage />} />
        </Route>
        <Route path="/demo" element={<DemoLayout />}>
          <Route index element={<DemoHomePage />} />
          <Route path="catalog" element={<CatalogPage />} />
          <Route
            path="resources"
            element={
              <DemoPage slug="resources">
                <ResourcesWidget />
              </DemoPage>
            }
          />
          <Route
            path="trainings"
            element={
              <DemoPage slug="trainings">
                <TrainingsWidget />
              </DemoPage>
            }
          />
          <Route
            path="governance"
            element={
              <DemoPage slug="governance">
                <GovernanceWidget />
              </DemoPage>
            }
          />
          <Route
            path="site-form"
            element={
              <DemoPage slug="site-form">
                <SiteFormWidget />
              </DemoPage>
            }
          />
          <Route
            path="maintenance-form"
            element={
              <DemoPage slug="maintenance-form">
                <MaintenanceFormWidget />
              </DemoPage>
            }
          />
          <Route
            path="completion"
            element={
              <DemoPage slug="completion">
                <CompletionWidget />
              </DemoPage>
            }
          />
          <Route
            path="kpi"
            element={
              <DemoPage slug="kpi">
                <KpiWidget />
              </DemoPage>
            }
          />
          <Route
            path="ot-kpi"
            element={
              <DemoPage slug="ot-kpi">
                <OtKpiWidget />
              </DemoPage>
            }
          />
          <Route
            path="page-kpi"
            element={
              <DemoPage slug="page-kpi">
                <PageKpiWidget />
              </DemoPage>
            }
          />
          <Route
            path="map"
            element={
              <DemoPage slug="map">
                <MapWidget />
              </DemoPage>
            }
          />
          <Route
            path="carousel"
            element={
              <DemoPage slug="carousel">
                <CarouselWidget />
              </DemoPage>
            }
          />
          <Route
            path="timeline"
            element={
              <DemoPage slug="timeline">
                <TimelineWidget />
              </DemoPage>
            }
          />
          <Route
            path="survey"
            element={
              <DemoPage slug="survey">
                <SurveyWidget />
              </DemoPage>
            }
          />
          <Route
            path="projects"
            element={
              <DemoPage slug="projects">
                <ProjectsWidget />
              </DemoPage>
            }
          />
          <Route
            path="confluence"
            element={
              <DemoPage slug="confluence">
                <ConfluenceWidget />
              </DemoPage>
            }
          />
          <Route
            path="engage"
            element={
              <DemoPage slug="engage">
                <EngageWidget />
              </DemoPage>
            }
          />
          <Route
            path="newsletter"
            element={
              <DemoPage slug="newsletter">
                <NewsletterWidget />
              </DemoPage>
            }
          />
        </Route>
      </Routes>
    </BrowserRouter>
  )
}
