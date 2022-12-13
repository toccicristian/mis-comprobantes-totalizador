class Orden_columnas:
    def __init__(self,pv,n_comp,t_comp,denominacion,n_documento,t_documento,t_cambio,neto,neto_no_g,exento,iva,total):
        self._pv=pv
        self._n_comp=n_comp
        self._t_comp=t_comp
        self._denominacion=denominacion
        self._n_documento=n_documento
        self._t_documento=t_documento
        self._t_cambio=t_cambio
        self._neto=neto
        self._neto_no_g=neto_no_g
        self._exento=exento
        self._iva=iva
        self._total=total

    @property
    def pv(self):
        return self._pv

    @pv.setter
    def pv(self, pv):
        self._pv=pv

    @property
    def n_comp(self):
        return self._n_comp

    @n_comp.setter
    def n_comp(self, n_comp):
        self._n_comp=n_comp

    @property
    def t_comp(self):
        return self._t_comp

    @t_comp.setter
    def t_comp(self, t_comp):
        self._t_comp=t_comp

    @property
    def denominacion(self):
        return self._denominacion

    @denominacion.setter
    def denominacion(self, denominacion):
        self._denominacion=denominacion

    @property
    def n_documento(self):
        return self._n_documento

    @n_documento.setter
    def n_documento(self,n_documento):
        self._n_documento=n_documento

    @property
    def t_documento(self):
        return self._t_documento

    @t_documento.setter
    def t_documento(self,t_documento):
        self._t_documento=t_documento

    @property
    def t_cambio(self):
        return self._t_cambio

    @t_cambio.setter
    def t_cambio(self, t_cambio):
        self._t_cambio=t_cambio

    @property
    def neto(self):
        return self._neto

    @neto.setter
    def neto(self, neto):
        self._neto=neto

    @property
    def neto_no_g(self):
        return self._neto_no_g

    @neto_no_g.setter
    def neto_no_g(self, neto_no_g):
        self._neto_no_g=neto_no_g

    @property
    def exento(self):
        return self._exento

    @exento.setter
    def exento(self, exento):
        self._exento=exento

    @property
    def iva(self):
        return self._iva

    @iva.setter
    def iva(self, iva):
        self._iva=iva

    @property
    def total(self):
        return self._total

    @total.setter
    def total(self,total):
        self._total=total
