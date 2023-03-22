import {
  createRouter,
  createWebHashHistory,
} from 'vue-router'

import Index from '../pages/index.vue'
import Excel from '../components/Excel.vue'
import BtnWrapper from '../components/BtnWrapper.vue'
import CodePen from '../components/CodePen.vue'
import Dragstart from '../components/Dragstart.vue'
import Hump from '../components/Hump.vue'
const routes = [{
  path: '/',
  component: Index,
  beforeEnter: (to, from) => {
    // reject the navigation
    // return false
  },
},
{
  path: '/excel',
  component: Excel,
},
{
  path: '/btnWrapper',
  component: BtnWrapper,
},
{
  path: '/codePen',
  component: CodePen,
},
{
  path: '/dragstart',
  component: Dragstart,
},
{
  path: '/hump',
  component: Hump,
},
]

const router = createRouter({
  history: createWebHashHistory(),
  routes,
})





export default router